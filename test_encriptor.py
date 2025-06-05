# Here begins a new world...

def text_to_binary(text):
    binary_list = []
    for char in text:
        # ord(char) → отримуємо числовий код символа
        # format(..., '08b') → перетворюємо в двійковий з 8 бітами
        binary_code = format(ord(char), '08b')
        binary_list.append(binary_code)
    return ' '.join(binary_list)

message = "My heart belongs to you 💛"
binary_message = text_to_binary(message)
print(binary_message)