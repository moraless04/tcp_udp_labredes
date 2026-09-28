#!/bin/sh
set -eu

exec /usr/bin/supervisord -n -c /etc/supervisor/conf.d/dns-lab.conf
