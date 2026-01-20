from datetime import datetime
import pandas as pd

from .config import AppConfig, get
from .file_lock import assert_files_closed
from .history_store import HistoryStore
from .logging_utils import make_logger
from .outlook_client import OutlookClient


def run_followup(cfg: AppConfig, month_override: str | None, display_only: bool) -> int:
    log = make_logger()

    history_path = get(cfg, "paths", "history_xlsx")
    history_sheet = get(cfg, "paths", "history_sheet", default="Auditoria De Estruturadas")

    month_ref = month_override or str(get(cfg, "followup", "month_reference", default="2026-01"))
    inbox_max = int(get(cfg, "outlook", "inbox_scan_max_items", default=5000))

    status_sent = str(get(cfg, "dispatch", "status_sent_label", default="Enviado"))
    label_replied = str(get(cfg, "followup", "mark_replied_label", default="Respondido"))
    label_reminded = str(get(cfg, "followup", "mark_reminded_label", default="Sem Retorno – Cobrado Novamente"))
    reminder_message = str(get(cfg, "followup", "reminder_message", default="Prezado(a),\n\nGentileza retornar.\n\nAtenciosamente,\nEquipe de Auditoria\n"))
    require_external_sender = bool(get(cfg, "followup", "require_external_sender", default=False))

    display_default = bool(get(cfg, "run_mode", "display_only_default", default=False))
    display_effective = bool(display_only or display_default)

    if not history_path:
        log.error("Config inválida: paths.history_xlsx é obrigatório.")
        return 2

    try:
        assert_files_closed([history_path])
    except Exception as e:
        log.error(str(e))
        return 2

    store = HistoryStore(history_path=history_path, sheet_name=history_sheet)
    df = store.load_history_df()

    for col in ["Data da Nova Cobrança", "Status Auditoria", "Data da Resposta", "Conteúdo da Resposta"]:
        if col not in df.columns:
            df[col] = None

    if "Data Envio" not in df.columns or "Status" not in df.columns:
        log.error('Histórico precisa ter colunas "Data Envio" e "Status".')
        return 2

    df["Data Envio"] = pd.to_datetime(df["Data Envio"], errors="coerce")

    df_filtrado = df[
        (df["Data Envio"].dt.strftime("%Y-%m") == month_ref)
        & (df["Status"].astype(str).str.lower() == status_sent.lower())
    ].copy()

    if df_filtrado.empty:
        log.info(f"Nenhum registro encontrado para mês={month_ref} e status={status_sent}.")
        return 0

    outlook = OutlookClient()

    my_email = None
    try:
        my_email = outlook.outlook.Session.CurrentUser.AddressEntry.GetExchangeUser().PrimarySmtpAddress.lower()
    except Exception:
        my_email = None

    for idx, row in df_filtrado.iterrows():
        nome_cliente = row.get("Nome do Cliente")
        entry_id = str(row.get("EntryID") or "").strip()

        if not entry_id:
            log.warn(f"Cliente {nome_cliente}: sem EntryID no histórico.")
            continue

        try:
            original = outlook.get_item_from_id(entry_id)
            conversation_id = getattr(original, "ConversationID", "")

            found, received_iso, body = outlook.scan_inbox_for_reply_by_conversation_id(conversation_id, max_items=inbox_max)

            if found and require_external_sender and my_email:
                try:
                    items = outlook.inbox_folder.Items
                    items.Sort("[ReceivedTime]", True)
                    validated = False
                    count = 0
                    for item in items:
                        count += 1
                        if count > inbox_max:
                            break
                        if getattr(item, "Class", None) != 43:
                            continue
                        if getattr(item, "ConversationID", "") != conversation_id:
                            continue
                        sender = str(getattr(item, "SenderEmailAddress", "") or "").lower()
                        if sender and sender != my_email:
                            received_iso = str(getattr(item, "ReceivedTime", None))
                            body = str(getattr(item, "Body", "") or "")
                            validated = True
                            break
                    found = validated
                except Exception:
                    pass

            if found:
                df.at[idx, "Status Auditoria"] = label_replied
                df.at[idx, "Data da Resposta"] = received_iso
                df.at[idx, "Conteúdo da Resposta"] = (body or "").strip()
                log.ok(f"[RESPONDIDO] Cliente {nome_cliente} – {received_iso or ''}")
            else:
                reply = original.Reply()
                reply.Body = reminder_message.format(nome_cliente=nome_cliente) + "\n\n" + reply.Body
                if display_effective:
                    reply.Display()
                else:
                    reply.Send()

                df.at[idx, "Data da Nova Cobrança"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                df.at[idx, "Status Auditoria"] = label_reminded
                df.at[idx, "Data da Resposta"] = None
                df.at[idx, "Conteúdo da Resposta"] = ""
                log.warn(f"[SEM RESPOSTA] Cliente {nome_cliente} – cobrança criada.")

        except Exception as e:
            log.error(f"Cliente {nome_cliente} – Falha ao processar follow-up: {e}")

    store.save_history_df(df)
    log.ok("Histórico atualizado com sucesso.")
    return 0
