import argparse
from pathlib import Path

from src.outlook_audit.config import load_config
from src.outlook_audit.dispatch import run_dispatch
from src.outlook_audit.followup import run_followup


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="outlook-audit-automation",
        description="Automação (sanitizada) de auditoria via Outlook: envio + rastreio + follow-up.",
    )
    p.add_argument(
        "--config",
        default="config.json",
        help="Caminho do config.json (padrão: config.json).",
    )

    sub = p.add_subparsers(dest="cmd", required=True)

    s1 = sub.add_parser("dispatch", help="Enviar e-mails de auditoria e registrar histórico (Excel).")
    s1.add_argument(
        "--dry-run",
        action="store_true",
        help="Não envia e-mails; apenas valida entradas e mostra o que seria enviado.",
    )
    s1.add_argument(
        "--display-only",
        action="store_true",
        help="Abre o e-mail (.Display) ao invés de enviar (.Send). Sobrescreve o config.",
    )

    s2 = sub.add_parser("followup", help="Verificar respostas e (opcionalmente) reenviar cobrança.")
    s2.add_argument(
        "--month",
        default=None,
        help="Mês de referência no formato YYYY-MM (sobrescreve followup.month_reference do config).",
    )
    s2.add_argument(
        "--display-only",
        action="store_true",
        help="Abre a cobrança (.Display) ao invés de enviar (.Send). Sobrescreve o config.",
    )

    return p


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    config_path = Path(args.config)
    if not config_path.exists():
        print(f"[ERRO] config.json não encontrado em: {config_path.resolve()}")
        print("      Copie o config.example.json para config.json e ajuste os caminhos/parâmetros.")
        return 2

    cfg = load_config(config_path)

    if args.cmd == "dispatch":
        return run_dispatch(cfg, dry_run=bool(args.dry_run), display_only=bool(args.display_only))

    if args.cmd == "followup":
        return run_followup(cfg, month_override=args.month, display_only=bool(args.display_only))

    print("[ERRO] Comando inválido.")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
