# -*- coding: utf-8 -*-
"""
Модуль для конвертации чисел в текст на русском языке
Поддержка валюты: тенге
"""

ONES = {
    0: '', 1: 'один', 2: 'два', 3: 'три', 4: 'четыре',
    5: 'пять', 6: 'шесть', 7: 'семь', 8: 'восемь', 9: 'девять',
    10: 'десять', 11: 'одиннадцать', 12: 'двенадцать', 13: 'тринадцать',
    14: 'четырнадцать', 15: 'пятнадцать', 16: 'шестнадцать',
    17: 'семнадцать', 18: 'восемнадцать', 19: 'девятнадцать'
}

ONES_FEMININE = {
    1: 'одна', 2: 'две'
}

TENS = {
    2: 'двадцать', 3: 'тридцать', 4: 'сорок', 5: 'пятьдесят',
    6: 'шестьдесят', 7: 'семьдесят', 8: 'восемьдесят', 9: 'девяносто'
}

HUNDREDS = {
    1: 'сто', 2: 'двести', 3: 'триста', 4: 'четыреста',
    5: 'пятьсот', 6: 'шестьсот', 7: 'семьсот', 8: 'восемьсот', 9: 'девятьсот'
}

# (единственное, 2-4, множественное, род: m/f)
UNITS = [
    ('', '', '', 'm'),
    ('тысяча', 'тысячи', 'тысяч', 'f'),
    ('миллион', 'миллиона', 'миллионов', 'm'),
    ('миллиард', 'миллиарда', 'миллиардов', 'm'),
    ('триллион', 'триллиона', 'триллионов', 'm'),
]


def get_plural_form(n, forms):
    """Получить правильную форму слова в зависимости от числа"""
    n = abs(n) % 100
    if 11 <= n <= 19:
        return forms[2]
    n = n % 10
    if n == 1:
        return forms[0]
    if 2 <= n <= 4:
        return forms[1]
    return forms[2]


def convert_group(n, feminine=False):
    """Конвертировать число от 0 до 999 в текст"""
    if n == 0:
        return ''

    result = []

    # Сотни
    hundreds = n // 100
    if hundreds:
        result.append(HUNDREDS[hundreds])

    # Десятки и единицы
    remainder = n % 100

    if remainder >= 20:
        tens = remainder // 10
        ones = remainder % 10
        result.append(TENS[tens])
        if ones:
            if feminine and ones in ONES_FEMININE:
                result.append(ONES_FEMININE[ones])
            else:
                result.append(ONES[ones])
    elif remainder > 0:
        if feminine and remainder in ONES_FEMININE:
            result.append(ONES_FEMININE[remainder])
        else:
            result.append(ONES[remainder])

    return ' '.join(result)


def number_to_text(n):
    """
    Конвертировать целое число в текст на русском языке

    Args:
        n: целое число (до триллионов)

    Returns:
        str: число прописью
    """
    if n == 0:
        return 'ноль'

    if n < 0:
        return 'минус ' + number_to_text(-n)

    groups = []
    group_index = 0

    while n > 0:
        group = n % 1000
        n //= 1000

        if group > 0:
            feminine = (group_index == 1)  # тысячи - женский род
            text = convert_group(group, feminine)

            if group_index > 0:
                unit = get_plural_form(group, UNITS[group_index][:3])
                text = f"{text} {unit}"

            groups.append(text)

        group_index += 1

    groups.reverse()
    return ' '.join(groups).strip()


def number_to_text_with_currency(n, currency='тенге'):
    """
    Конвертировать число в текст с указанием валюты

    Args:
        n: число (может быть float, копейки будут отброшены)
        currency: название валюты

    Returns:
        str: число прописью с валютой
    """
    n = int(n)
    text = number_to_text(n)
    return f"{text} {currency}"


def format_number_with_text(n):
    """
    Форматировать число с текстом в скобках
    Пример: 7 652 278 (семь миллионов шестьсот пятьдесят две тысячи двести семьдесят восемь)

    Args:
        n: число

    Returns:
        str: форматированная строка
    """
    n = int(n)
    formatted_num = f"{n:,}".replace(',', ' ')
    text = number_to_text(n)
    return f"{formatted_num} ({text})"


if __name__ == '__main__':
    # Тесты
    test_numbers = [
        0, 1, 2, 11, 21, 100, 101, 111, 121, 200,
        1000, 1001, 2000, 5000, 21000,
        1000000, 7652278, 6551320, 799832, 301126
    ]

    print("Тестирование перевода чисел прописью:\n")
    for num in test_numbers:
        print(f"{num:>12,} -> {number_to_text(num)}")

    print("\n\nПримеры из справки:")
    print(format_number_with_text(7652278))
    print(format_number_with_text(6551320))
    print(format_number_with_text(799832))
    print(format_number_with_text(301126))
