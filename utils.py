def text_to_binary(ENCODED_INFORMATION):
    string = bytes.decode(str.encode(ENCODED_INFORMATION))
    bits_arr = list(map(int, ''.join([bin(ord(i)).lstrip('0b').rjust(8,'0') for i in string])))
    return ''.join(str(e) for e in bits_arr)