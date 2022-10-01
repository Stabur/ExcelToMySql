import uuid
import hashlib

def stabur_cripto(password):
    # Модуль для хэширования с последующим разделением на части и переменой мест этих частей
    password_enc = password.encode('utf-8')
    #password_dec = password_enc.decode('utf-8')
    #print("Пароль : " + password_dec)
    hash_var = hashlib.sha256(password_enc).hexdigest()
    #print("Хэш sha256 : " + hash_var)
    #hash_var_len = len(hash_var)
    #print("Кол-во символов в хэше : " + str(hash_var_len))
    s1 = hash_var[:len(hash_var) // 2]
    s2 = hash_var[len(hash_var) // 2:]
    ss1 = s1[:len(s1) // 2]
    ss2 = s1[len(s1) // 2:]
    ss3 = s2[:len(s2) // 2]
    ss4 = s2[len(s2) // 2:]
    hash_sha256_mix = (ss4 + ss1 + ss3 + ss2)
    # uuid используется для генерации случайного числа для соли(salt)
    salt1 = uuid.uuid4().hex
    salt2 = uuid.uuid4().hex
    salt3 = uuid.uuid4().hex
    super_hash = hashlib.sha256(hash_sha256_mix.encode()).hexdigest()
    return salt2 + super_hash + salt3 + salt1 # Перемешиваем хэш с солью, получаем 160-и значный ключ
