import configparser
import xlwt
import csv
import time
import paramiko
import logging
from scp import SCPClient
from io import StringIO

config = configparser.ConfigParser()
config.read('convert.ini')
base_config = config['lora']
logging_config = config['logging']
ssh_config = config['ssh']

logging.basicConfig(
    filename=logging_config['log_file'],
    level=int(logging_config['log_level'])
)


def createSSHClient(server, port, user, password):
    client = paramiko.SSHClient()
    client.load_system_host_keys()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    client.connect(server, port, user, password)
    return client


ssh = scp = None

while True:
    time.sleep(float(base_config['trigger']) / 1000)
    if not ssh and ssh_config['ssh_host']:
        try:
            ssh = createSSHClient(
                ssh_config['ssh_host'],
                int(ssh_config['ssh_port']),
                ssh_config['ssh_user'],
                ssh_config['ssh_pass']
            )
        except Exception as e:
            logging.error(e)
            ssh = scp = None
            continue
    if not scp and ssh:
        try:
            scp = SCPClient(ssh.get_transport())
        except Exception as e:
            logging.error(e)
            ssh.close()
            ssh = scp = None
            continue

    try:
        if scp:
            try:
                scp.get(ssh_config['ssh_file'], base_config['input'])
            except Exception as e:
                logging.error(e)
                ssh = scp = None
        with open(base_config['input'], 'rb') as csvfile:
            xlsout = []
            csvfile = StringIO(
                csvfile.read().decode(errors="ignore").replace('\x00', '')
            )
            workbook = xlwt.Workbook()
            worksheet = workbook.add_sheet('Mappe 1')
            vabreader = csv.reader(csvfile, delimiter=';')

            csvin = {}

            for i, row in enumerate(vabreader):
                row = [r.replace('?', '').strip() for r in row]
                if row:
                    csvin[row[0]] = row[1:]

            if base_config['mapping']:
                mapping = base_config['mapping'].split()
            else:
                mapping = csvin.keys()

            for i, key in enumerate(mapping):
                worksheet.write(i, 0, key)
                if key in csvin:
                    for j, cell in enumerate(csvin[key]):
                        worksheet.write(i, j+1, cell)
                else:
                    for j in range(0, len(list(csvin.values())[0])):
                        worksheet.write(i, j+1, '0')

            with open(base_config['output'], 'bw') as f:
                workbook.save(f)
    except Exception as e:
        logging.error(e)
