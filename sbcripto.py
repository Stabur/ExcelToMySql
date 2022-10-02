import uuid
import hashlib

def stabur_cripto(password):
    # Модуль для хэширования с последующим разделением на части и переменой мест этих частей
    password_enc = password.encode('utf-8')
    hash_var = hashlib.sha256(password_enc).hexdigest()
    s1 = hash_var[:len(hash_var) // 2]
    s2 = hash_var[len(hash_var) // 2:]
    ss1 = s1[:len(s1) // 2]
    ss2 = s1[len(s1) // 2:]
    ss3 = s2[:len(s2) // 2]
    ss4 = s2[len(s2) // 2:]
    hash_sha256_mix = (ss4 + ss1 + ss3 + ss2)
    salt1 = uuid.uuid4().hex
    salt2 = uuid.uuid4().hex
    salt3 = uuid.uuid4().hex
    super_hash = hashlib.sha256(hash_sha256_mix.encode()).hexdigest()
    return salt2 + super_hash + salt3 + salt1
