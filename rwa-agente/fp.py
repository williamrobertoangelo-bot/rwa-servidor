import socket, getpass, platform, uuid, hashlib, subprocess

def serial():
    try:
        r = subprocess.check_output('wmic diskdrive get serialnumber', shell=True, stderr=subprocess.DEVNULL)
        return r.decode(errors='ignore').split()[-1]
    except: return ''

def uuid_bios():
    try:
        r = subprocess.check_output('wmic csproduct get uuid', shell=True, stderr=subprocess.DEVNULL)
        return r.decode(errors='ignore').split()[-1]
    except: return ''

base = '|'.join([socket.gethostname(), getpass.getuser(), platform.platform(), str(uuid.getnode()), serial(), uuid_bios()])
print(hashlib.sha256(base.encode()).hexdigest().upper()[:32])