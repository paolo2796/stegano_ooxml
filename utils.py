def text_to_binary(ENCODED_INFORMATION):
    a_byte_array = bytearray(ENCODED_INFORMATION, "utf8")
    byte_list = []
    for byte in a_byte_array:
        binary_representation = bin(byte)
        byte_list.append(binary_representation)
    ENCODED_INFORMATION_BITS = ''.join(byte_list).replace("0b", "0")
    return ENCODED_INFORMATION_BITS