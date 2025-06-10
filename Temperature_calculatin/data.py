# Ця програма входить до проєкту Temperature_calculating
# Відповідає за завантаження та первинну обробку даних

def load_data(filename: str) -> list[str]:
    with open(filename, "r") as file:
        return file.readlines()
    
def clean_data(temp_data: list[str]) -> list[float]:
    return [float(temp.strip()) for temp in temp_data if temp.strip()]

