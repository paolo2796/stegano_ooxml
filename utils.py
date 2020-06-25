def binarytoDecimal(binary, i=0):
    # If we reached last character
    n = len(binary)
    if (i == n - 1):
        return int(binary[i]) - 0
    # Add current tern and recur for
    # remaining terms
    return (((int(binary[i]) - 0) << (n - i - 1)) +
            binarytoDecimal(binary, i + 1))