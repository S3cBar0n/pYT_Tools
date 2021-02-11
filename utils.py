# import platform
# import subprocess

import socket

# Older Ping command, replaced by the lookup function to perform a similar action
# def ping(host):
#     parameter = '-n' if platform.system().lower() == 'windows' else '-c'
#     ping_command = ['ping', parameter, '1', host]
#     response = subprocess.Popen(ping_command)
#
#     if response == 0:
#         return print(f'{host} is online...')
#     else:
#         return print(f'{host} is offline...')


def lookup(host):
    try:
        host_info = str(socket.getaddrinfo(host, 0, socket.AF_INET, 0, 0))

        for sym in (
        ('(', ''), (')', ''), ("<AddressFamily.AF_INET: 2>, 0, 0, '', ", ""), ('[', ''), (']', ''), (" ", ""),
        ("'", ""),
        (",0", "")):
            host_info = host_info.replace(*sym)
        host_info = host_info.split(",")
        for ip_address in host_info:
            print(ip_address)

    except:
        print('Could not locate the device on the network...')
