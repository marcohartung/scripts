#!/bin/sh
# Do the system update (Ubuntu)
#
# modify /etc/crontab to run automated (daily)
# 0 0 * * *    root   /usr/local/bin/do_update.sh
#
# Marco Hartung, 30.12.2020
# Version 1.0.0

apt-get update
apt-get upgrade -y
apt-get dist-upgrade -y
apt-get autoremove -y
