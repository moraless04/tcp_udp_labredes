#!/usr/bin/env bash
set -Eeuo pipefail

if [[ "$(id -u)" -ne 0 ]]; then
  echo "Execute este script como root dentro do container monitor." >&2
  exit 1
fi

if ! command -v apt-get >/dev/null 2>&1; then
  echo "Este script requer apt-get (imagem Debian/Ubuntu)." >&2
  exit 1
fi

echo "Atualizando a lista de pacotes..."
apt-get update

echo "Instalando ferramentas para os experimentos..."
DEBIAN_FRONTEND=noninteractive apt-get install -y --no-install-recommends \
  ca-certificates \
  curl \
  dnsutils \
  ethtool \
  iproute2 \
  net-tools \
  procps \
  python3

rm -rf /var/lib/apt/lists/*

echo
echo "Ferramentas disponíveis:"
for tool in curl dig nslookup ethtool ip ss netstat ps python3; do
  if command -v "$tool" >/dev/null 2>&1; then
    printf '  %-10s %s\n' "$tool" "$(command -v "$tool")"
  else
    echo "  $tool não foi encontrado após a instalação." >&2
    exit 1
  fi
done

echo
echo "Preparação do monitor concluída."
