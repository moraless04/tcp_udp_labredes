#!/bin/sh
set -eu

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
LAB_DIR=$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)

NETWORK=dns-lab-net
SUBNET=10.53.0.0/24
GATEWAY=10.53.0.1
STATION_IP=10.53.0.10
RESOLVER_IP=10.53.0.53
POD=dns-station
CLIENT=dns-client
MONITOR=dns-monitor
RESOLVER=dns-resolver

CLIENT_IMAGE=dns-lab-client:local
MONITOR_IMAGE=dns-lab-monitor:local
RESOLVER_IMAGE=dns-lab-resolver:local
NOVNC_URL=http://127.0.0.1:6080/vnc.html

log() {
    printf '%s\n' "$*"
}

die() {
    printf 'Erro: %s\n' "$*" >&2
    exit 1
}

require_podman() {
    command -v podman >/dev/null 2>&1 || die "Podman não foi encontrado no PATH."
    podman info >/dev/null 2>&1 || die "Podman Machine não está acessível. Inicie-a com: podman machine start"
}

container_exists() {
    podman container exists "$1"
}

pod_exists() {
    podman pod exists "$1"
}

image_exists() {
    podman image exists "$1"
}

network_exists() {
    podman network exists "$1"
}

build_images() {
    require_podman
    log "Construindo imagem do cliente..."
    podman build -t "$CLIENT_IMAGE" -f "$LAB_DIR/Containerfile.client" "$LAB_DIR"
    log "Construindo imagem do monitor..."
    podman build -t "$MONITOR_IMAGE" -f "$LAB_DIR/Containerfile.monitor" "$LAB_DIR"
    log "Construindo imagem do resolvedor..."
    podman build -t "$RESOLVER_IMAGE" -f "$LAB_DIR/Containerfile.resolver" "$LAB_DIR"
}

ensure_images() {
    if ! image_exists "$CLIENT_IMAGE" || ! image_exists "$MONITOR_IMAGE" || ! image_exists "$RESOLVER_IMAGE"; then
        build_images
    fi
}

ensure_network() {
    if ! network_exists "$NETWORK"; then
        log "Criando rede $NETWORK ($SUBNET)..."
        podman network create --disable-dns --subnet "$SUBNET" --gateway "$GATEWAY" "$NETWORK" >/dev/null
    fi
}

wait_for_resolver() {
    attempts=0
    while [ "$attempts" -lt 30 ]; do
        if podman exec "$RESOLVER" dig +time=1 +tries=1 @127.0.0.1 example.org A >/dev/null 2>&1; then
            return 0
        fi
        attempts=$((attempts + 1))
        sleep 1
    done
    podman logs "$RESOLVER" >&2 || true
    die "O resolvedor não ficou pronto dentro do tempo esperado."
}

wait_for_novnc() {
    attempts=0
    while [ "$attempts" -lt 45 ]; do
        if curl -fsS "$NOVNC_URL" >/dev/null 2>&1; then
            return 0
        fi
        attempts=$((attempts + 1))
        sleep 1
    done
    podman logs "$MONITOR" >&2 || true
    die "O noVNC não ficou disponível em $NOVNC_URL."
}

create_resolver() {
    if ! container_exists "$RESOLVER"; then
        log "Iniciando resolvedor em $RESOLVER_IP..."
        podman run -d \
            --name "$RESOLVER" \
            --network "$NETWORK" \
            --ip "$RESOLVER_IP" \
            "$RESOLVER_IMAGE" >/dev/null
    else
        podman start "$RESOLVER" >/dev/null
    fi
    wait_for_resolver
}

create_station() {
    if ! pod_exists "$POD"; then
        log "Criando pod da estação em $STATION_IP..."
        podman pod create \
            --name "$POD" \
            --network "$NETWORK" \
            --ip "$STATION_IP" \
            --dns "$RESOLVER_IP" \
            -p 127.0.0.1:6080:6080 >/dev/null
    fi

    mkdir -p "$LAB_DIR/captures"
    chmod a+rwx "$LAB_DIR/captures"

    if ! container_exists "$MONITOR"; then
        log "Iniciando monitor Wireshark/noVNC..."
        podman run -d \
            --name "$MONITOR" \
            --pod "$POD" \
            --cap-add NET_ADMIN \
            --cap-add NET_RAW \
            -v "$LAB_DIR/captures:/captures:rw" \
            "$MONITOR_IMAGE" >/dev/null
    fi

    if ! container_exists "$CLIENT"; then
        log "Iniciando cliente DNS..."
        podman run -d \
            --name "$CLIENT" \
            --pod "$POD" \
            "$CLIENT_IMAGE" >/dev/null
    fi

    podman pod start "$POD" >/dev/null
    wait_for_novnc
}

up() {
    require_podman
    ensure_images
    ensure_network
    create_resolver
    create_station
    log ""
    log "Ambiente pronto."
    log "Wireshark/noVNC: $NOVNC_URL"
    log "Terminal do aluno: ./scripts/lab.sh shell"
}

start_lab() {
    require_podman
    if ! pod_exists "$POD" || ! container_exists "$RESOLVER"; then
        up
        return
    fi
    podman start "$RESOLVER" >/dev/null
    podman pod start "$POD" >/dev/null
    wait_for_resolver
    wait_for_novnc
    log "Ambiente iniciado: $NOVNC_URL"
}

stop_lab() {
    require_podman
    if pod_exists "$POD"; then
        podman pod stop "$POD" >/dev/null
    fi
    if container_exists "$RESOLVER"; then
        podman stop "$RESOLVER" >/dev/null
    fi
    log "Ambiente parado; contêineres e capturas foram preservados."
}

down() {
    require_podman
    if pod_exists "$POD"; then
        podman pod rm -f "$POD" >/dev/null
    fi
    if container_exists "$RESOLVER"; then
        podman rm -f "$RESOLVER" >/dev/null
    fi
    log "Pod e contêineres removidos; imagens, rede e capturas foram preservadas."
}

clean() {
    require_podman
    down
    if network_exists "$NETWORK"; then
        podman network rm "$NETWORK" >/dev/null
    fi
    for image in "$CLIENT_IMAGE" "$MONITOR_IMAGE" "$RESOLVER_IMAGE"; do
        if image_exists "$image"; then
            podman image rm "$image" >/dev/null
        fi
    done
    log "Ambiente, rede e imagens locais removidos. O diretório captures/ foi preservado."
}

status() {
    require_podman
    log "Pod:"
    podman pod ps --filter "name=$POD"
    log ""
    log "Contêineres:"
    podman ps -a --filter "name=$CLIENT" --filter "name=$MONITOR" --filter "name=$RESOLVER"
    log ""
    log "Acesso: $NOVNC_URL"
}

shell_client() {
    require_podman
    container_exists "$CLIENT" || die "O cliente não existe. Execute: ./scripts/lab.sh up"
    podman exec -it "$CLIENT" bash
}

flush_dns() {
    require_podman
    container_exists "$RESOLVER" || die "O resolvedor não existe. Execute: ./scripts/lab.sh up"
    podman exec "$RESOLVER" unbound-control -c /etc/unbound/unbound.conf flush_zone .
}

smoke_test() {
    require_podman
    container_exists "$CLIENT" || die "O cliente não existe. Execute: ./scripts/lab.sh up"
    log "Servidor configurado no cliente:"
    podman exec "$CLIENT" cat /etc/resolv.conf
    log ""
    log "Consulta pelo resolvedor do laboratório:"
    podman exec "$CLIENT" dig +short www.google.com A
    log ""
    log "Consulta direta ao Google Public DNS:"
    podman exec "$CLIENT" dig +time=3 +tries=1 +short @8.8.8.8 www.wikipedia.org A
    log ""
    log "Consulta direta ao Cloudflare DNS:"
    podman exec "$CLIENT" dig +time=3 +tries=1 +short @1.1.1.1 www.wikipedia.org A
    log ""
    log "Teste concluído. Confirme no Wireshark os destinos e as respostas."
}

doctor() {
    require_podman
    log "Podman: OK"
    if container_exists "$CLIENT"; then
        podman exec "$CLIENT" dig +time=3 +tries=1 @8.8.8.8 example.org A >/dev/null \
            && log "UDP/53 para 8.8.8.8: OK" \
            || log "UDP/53 para 8.8.8.8: FALHOU"
        podman exec "$CLIENT" dig +time=3 +tries=1 @1.1.1.1 example.org A >/dev/null \
            && log "UDP/53 para 1.1.1.1: OK" \
            || log "UDP/53 para 1.1.1.1: FALHOU"
        curl -fsS "$NOVNC_URL" >/dev/null \
            && log "noVNC: OK ($NOVNC_URL)" \
            || log "noVNC: FALHOU ($NOVNC_URL)"
    else
        log "Ambiente ainda não iniciado; execute: ./scripts/lab.sh up"
    fi
}

usage() {
    cat <<'EOF'
Uso: ./scripts/lab.sh COMANDO

Comandos:
  build       constrói as três imagens
  up          cria/inicia todo o ambiente
  start       reinicia um ambiente previamente parado
  stop        para e preserva o ambiente
  down        remove pod e contêineres; preserva imagens, rede e capturas
  clean       remove ambiente, rede e imagens; preserva captures/
  status      mostra o estado do laboratório
  shell       abre um terminal no dns-client
  flush-dns   limpa o cache do Unbound
  test        executa consultas de verificação
  doctor      verifica Podman, DNS público e noVNC
EOF
}

command=${1:-help}
case "$command" in
    build) build_images ;;
    up) up ;;
    start) start_lab ;;
    stop) stop_lab ;;
    down) down ;;
    clean) clean ;;
    status) status ;;
    shell) shell_client ;;
    flush-dns) flush_dns ;;
    test) smoke_test ;;
    doctor) doctor ;;
    help|-h|--help) usage ;;
    *) usage >&2; exit 2 ;;
esac
