# Here begins a new world...

import math
from log import log_info, log_warning, log_error

def calculate_squere_root(numbers: list) -> None:
    for num in numbers:
        try:
            if num < 0:
                log_warning(f"Знайдено від'ємне число: {num}. Пропускаємо.")
                continue

            root = math.sqrt(num)
            log_info(f"Квадратний корінь з {num} = {root:.2f}")
        
        except Exception as e:
            log_error(f"Помилка при обчисленні кореня для {num}: {e}")

if __name__ == "__main__":
    numbers = [16, -4, 9, 25, 0, 4, "16"]
    calculate_squere_root(numbers)