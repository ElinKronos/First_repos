# Here begins a new world...

def binary_to_text(binary):
    chars = binary.split()  # розділяємо рядок на байти
    return ''.join(chr(int(b, 2)) for b in chars)

binary_message = "01010100 01110111 01101111 00100000 01101000 01100101 01100001 01110010 01110100 01110011 00100000 01101111 01101110 01100101 00100000 01100010 01100101 01100001 01110100"
print(binary_to_text(binary_message))
