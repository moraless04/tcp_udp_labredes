# Preparação do monitor para a atividade TCP/UDP

Este guia prepara o container `monitor` para executar os experimentos de
[`monitor_tcp_udp.pdf`](monitor_tcp_udp.pdf) como cliente e capturar o próprio
tráfego na interface `eth0`.

## 1. Inicie o laboratório em Linux

Execute os comandos em uma máquina Linux com Podman e o comando `podman compose`
disponíveis. No diretório `_labredes`, suba o laboratório:

```bash
./scripts/lab.sh up
```

Não é necessário iniciar uma máquina virtual do Podman. Se o laboratório já
estiver ativo, não é necessário subi-lo novamente.

## 2. Desative o espelhamento dos nós

Como os comandos do experimento serão executados dentro do próprio monitor,
ele captura as requisições e respostas da sua `eth0` diretamente. Para evitar
que tráfego dos outros containers apareça junto na captura, desative o
espelhamento. No diretório `_labredes`, execute:

```bash
./scripts/lab.sh mirror remove
```

## 3. Copie e execute o atualizador

Abra um terminal no diretório que contém este README e `update_lab.sh`. Copie o
script para o container e execute-o lá:

```bash
podman cp update_lab.sh monitor:/tmp/update_lab.sh
podman exec -it monitor bash /tmp/update_lab.sh
```

O container `monitor` executa como root por padrão. O script atualiza a lista
de pacotes e instala:

- `curl` e certificados CA, para requisições HTTP e HTTPS;
- `dnsutils`, que fornece `dig` e `nslookup`;
- `python3`, para gerar datagramas UDP;
- `ethtool`, para consultar recursos de checksum da interface;
- `iproute2`, `net-tools` e `procps`, para comandos como `ip`, `ss`, `netstat`
  e `ps`.

É necessário que o monitor tenha acesso à Internet durante a instalação, para
baixar os pacotes. Ao final, o script lista os comandos instalados ou
disponíveis.

## 4. Capture e execute os experimentos

No diretório `_labredes`, inicie uma captura salva em `resultados/` e abra um
terminal no monitor:

```bash
./scripts/lab.sh capture start monitor-tcp-udp.pcapng
./scripts/lab.sh shell monitor
```

Execute os comandos Linux do experimento dentro desse terminal. Por exemplo:

```bash
dig @8.8.8.8 www.pucrs.br A
curl -s -o /dev/null http://example.com/
```

Ao terminar, saia do shell do monitor e encerre a captura pelo terminal do
host:

```bash
exit
./scripts/lab.sh capture stop
```

Também é possível acompanhar a captura ao vivo pelo Wireshark no NoVNC em
`http://127.0.0.1:6080/`.

## Observações

- A instalação feita por `update_lab.sh` vale para o container atual. Se o monitor for removido e criado novamente, execute o script outra vez. Para  manter as ferramentas em todas as recriações, elas devem ser incluídas na  imagem do monitor.

