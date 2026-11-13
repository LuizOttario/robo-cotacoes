#!/usr/bin/env bash
# Agenda o robô pra rodar sozinho todo dia (Linux com systemd, ex.: CachyOS, Ubuntu, Fedora).
#
#   ./agendar.sh          -> agenda pra rodar todo dia às 9h
#   ./agendar.sh 18 30    -> agenda pra rodar todo dia às 18h30
#   ./agendar.sh remover  -> tira o agendamento
#   ./agendar.sh status   -> mostra quando vai ser a próxima execução
#
# Se o computador estiver desligado na hora marcada, o robô roda assim que ligar.
# O que o robô imprimir fica salvo em robo.log, na mesma pasta.

set -e
PASTA="$(cd "$(dirname "$0")" && pwd)"
PYTHON="$PASTA/.venv/bin/python"
DESTINO="$HOME/.config/systemd/user"
NOME="robo-cotacoes"

case "$1" in
    remover)
        systemctl --user disable --now "$NOME.timer" 2>/dev/null || true
        rm -f "$DESTINO/$NOME.service" "$DESTINO/$NOME.timer"
        systemctl --user daemon-reload
        echo "Agendamento removido."
        exit 0 ;;
    status)
        systemctl --user list-timers "$NOME.timer"
        exit 0 ;;
esac

if [ ! -x "$PYTHON" ]; then
    echo "Não achei o Python em $PYTHON"
    echo "Rode antes: python -m venv .venv && .venv/bin/pip install -r requirements.txt"
    exit 1
fi

HORA="${1:-9}"
MINUTO="${2:-0}"
mkdir -p "$DESTINO"

# O "serviço" diz O QUE rodar
cat > "$DESTINO/$NOME.service" <<EOF
[Unit]
Description=Robô de cotações (Python + RPA + Excel)
After=network-online.target

[Service]
Type=oneshot
WorkingDirectory=$PASTA
ExecStart=/bin/sh -c '"$PYTHON" robo_cotacoes.py >> robo.log 2>&1'
EOF

# O "timer" diz QUANDO rodar
cat > "$DESTINO/$NOME.timer" <<EOF
[Unit]
Description=Roda o robô de cotações todo dia

[Timer]
OnCalendar=*-*-* $(printf "%02d:%02d" "$HORA" "$MINUTO"):00
Persistent=true

[Install]
WantedBy=timers.target
EOF

systemctl --user daemon-reload
systemctl --user enable --now "$NOME.timer"
printf "Pronto! O robô vai rodar todo dia às %02d:%02d.\n" "$HORA" "$MINUTO"
echo "Pra testar agora: systemctl --user start $NOME.service && tail robo.log"
