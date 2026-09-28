# Laboratório de monitoração DNS com Podman

Este ambiente executa os experimentos do `monitor_dns_v2.docx` em três contêineres:

- `dns-client`: terminal com `nslookup` e `dig`;
- `dns-monitor`: Wireshark gráfico acessível por noVNC;
- `dns-resolver`: Unbound usado como resolvedor padrão controlável.

`dns-client` e `dns-monitor` pertencem ao mesmo pod e compartilham a interface de rede. Por isso, o Wireshark consegue observar as consultas geradas no terminal do cliente.

## Requisitos

- Podman 5 ou superior;
- Podman Machine iniciada no macOS ou Windows;
- navegador com acesso a `127.0.0.1:6080`;
- acesso externo a DNS UDP/TCP porta 53.

## Uso rápido

```sh
./scripts/lab.sh up
```

Abra <http://127.0.0.1:6080/vnc.html>, clique em **Connect** e confirme que o Wireshark está capturando na interface `eth0`. O acesso não solicita senha.

Em outro terminal:

```sh
./scripts/lab.sh shell
nslookup www.google.com
```

Para limpar o cache do resolvedor antes de uma nova rodada:

```sh
./scripts/lab.sh flush-dns
```

Para executar a verificação automatizada:

```sh
./scripts/lab.sh test
./scripts/lab.sh doctor
```

## Encerramento

```sh
./scripts/lab.sh stop   # preserva contêineres
./scripts/lab.sh down   # remove pod e contêineres
./scripts/lab.sh clean  # remove também rede e imagens locais
```

Os três comandos preservam `captures/`. A remoção das capturas deve ser feita explicitamente pelo usuário.

## Segurança e acesso

A porta noVNC é publicada somente em `127.0.0.1`, por isso o acesso local não exige senha. A porta VNC interna não é exposta. Não altere o mapeamento para `0.0.0.0:6080` sem adicionar autenticação e controles de acesso. O monitor recebe apenas as capacidades `NET_RAW` e `NET_ADMIN`, necessárias para a captura, sem usar `--privileged` ou a rede do host.

## Diagnóstico

Se o laboratório não iniciar, verifique:

```sh
podman machine start
./scripts/lab.sh doctor
./scripts/lab.sh status
podman logs dns-monitor
podman logs dns-resolver
```

VPNs e redes institucionais podem bloquear consultas diretas a `8.8.8.8` e `1.1.1.1`. Nesse caso, a Parte D do exercício requer liberação de UDP/TCP porta 53.
