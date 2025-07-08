# vigenere cipher
# https://stackoverflow.com/a/2490718/1675586

import base64

def encode(key, string) -> str:
    encoded_chars = []
    for i in range(len(string)):
        key_c = key[i % len(key)]
        encoded_c = chr(ord(string[i]) + ord(key_c) % 256)
        encoded_chars.append(encoded_c)
    encoded_string = ''.join(encoded_chars)
    encoded_string = encoded_string.encode('latin')
    b64 = base64.urlsafe_b64encode(encoded_string).rstrip(b'=')
    return b64.decode('latin')

def decode(key, string) -> str:
    string = string.encode('latin')
    string = base64.urlsafe_b64decode(string + b'===')
    string = string.decode('latin')
    encoded_chars = []
    for i in range(len(string)):
        key_c = key[i % len(key)]
        encoded_c = chr((ord(string[i]) - ord(key_c) + 256) % 256)
        encoded_chars.append(encoded_c)
    encoded_string = ''.join(encoded_chars)
    return encoded_string