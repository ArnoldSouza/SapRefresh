# -*- coding: utf-8 -*-
"""
Created on 3/8/2021
Author: Arnold Souza
Email: arnoldporto@gmail.com
"""
import os

from cryptography.fernet import Fernet


def secret_encode(string):
    """Encrypt the secret string"""
    key = os.environ.get('SAP_KEY').encode('utf-8')
    f = Fernet(key)
    token = f.encrypt(string.encode('utf-8'))
    return token


def secret_decode(token):
    """Decrypt the secret string"""
    key = os.environ.get('SAP_KEY').encode('utf-8')
    f = Fernet(key)
    value = f.decrypt(token.encode('utf-8')).decode('utf-8')
    return value


if __name__ == '__main__':
    string_test_encode = 'message_secret'
    secret_string = secret_encode(string_test_encode)
    print(secret_string)
