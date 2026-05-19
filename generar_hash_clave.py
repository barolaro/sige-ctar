import hashlib
import getpass

password = getpass.getpass("Ingrese clave a convertir en hash: ")
print(hashlib.sha256(password.encode("utf-8")).hexdigest())
