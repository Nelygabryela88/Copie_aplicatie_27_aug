import socket
s = socket.socket()
try:
    s.connect(('smtp.vitesco-technologies.net', 587))
    print("CONNECTION SUCCESSFUL")
except Exception as e:
    print("CONNECTION FAILED:", e)
s.close()