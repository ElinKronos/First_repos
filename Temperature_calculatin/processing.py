# Ця програма входить до проєкту Temperature_Calculating
# Вона обробляє дані: обчислення середньої температури, мінімальної, максимальної та медіанної.

def calc_statistics(temperatures: list[float]) -> dict:
    if not temperatures:
        return None
    
    min_temp = min(temperatures)
    max_temp = max(temperatures)
    avg_temp = sum(temperatures) / len(temperatures)
    median_temp = calc_median(temperatures)

    return {
        "min": min_temp,
        "max": max_temp,
        "average": avg_temp,
        "median": median_temp,
    }

def calc_median(temperatures: list[float]) -> float:
    temperatures.sort()
    n = len(temperatures)
    mid = n // 2
    if n % 2 == 0:
        return (temperatures[mid - 1] + temperatures[mid]) / 2
    else:
        return temperatures[mid]
