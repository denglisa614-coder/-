# -*- coding: utf-8 -*-
"""金额处理：人民币大写、数值格式化。"""
from decimal import Decimal, ROUND_HALF_UP

_DIGITS = "零壹贰叁肆伍陆柒捌玖"
_UNITS = ["", "拾", "佰", "仟"]
_BIG = ["", "万", "亿", "兆"]


def rmb_upper(amount) -> str:
    """把金额转成标准人民币大写，如 198.33 -> 壹佰玖拾捌元叁角叁分。

    处理规则：
    - 无角无分 -> 加「整」，如 198 -> 壹佰玖拾捌元整
    - 有角无分 -> 角后不加整，如 198.3 -> 壹佰玖拾捌元叁角
    - 有分 -> 正常，如 198.33 -> 壹佰玖拾捌元叁角叁分
    """
    n = Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if n < 0:
        raise ValueError("金额不能为负")
    integer = int(n)
    frac = int((n * 100).to_integral_value(rounding=ROUND_HALF_UP)) % 100
    jiao, fen = frac // 10, frac % 10

    # 整数部分转大写
    def _int_upper(num: int) -> str:
        if num == 0:
            return "零"
        groups = []
        while num > 0:
            groups.append(num % 10000)
            num //= 10000
        parts = []
        for gi in range(len(groups) - 1, -1, -1):
            g = groups[gi]
            s = _group_upper(g)
            if s:
                # 若该组 < 1000 且前面已有更高位组，需要补「零」
                if gi < len(groups) - 1 and g < 1000:
                    parts.append("零")
                parts.append(s + _BIG[gi])
        result = "".join(parts)
        # 去掉可能多余的前导零（例如 100000001 中的连续零）
        while "零零" in result:
            result = result.replace("零零", "零")
        return result.rstrip("零") if result.endswith("零") and result != "零" else result

    def _group_upper(g: int) -> str:
        if g == 0:
            return ""
        s = ""
        unit_pos = 0
        zero_pending = False
        while g > 0:
            d = g % 10
            if d == 0:
                zero_pending = True
            else:
                if zero_pending and s:
                    s = "零" + s
                zero_pending = False
                s = _DIGITS[d] + _UNITS[unit_pos] + s
            g //= 10
            unit_pos += 1
        return s

    int_str = _int_upper(integer)
    if integer == 0:
        int_str = "零"

    if jiao == 0 and fen == 0:
        return int_str + "元整"
    if fen == 0:
        return int_str + "元" + _DIGITS[jiao] + "角"
    return int_str + "元" + _DIGITS[jiao] + "角" + _DIGITS[fen] + "分"


def fmt_money(amount) -> str:
    """把金额转成去掉多余小数位的字符串：39.80->39.8, 59.00->59, 4.59->4.59。"""
    d = Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    if d == d.to_integral_value():
        return str(int(d))
    s = format(d, "f")
    if s.endswith("0"):
        s = s.rstrip("0")
    return s


def parse_money(s: str):
    """从字符串里解析金额，容错：全角/半角、千分位逗号、¥￥、多余空格。"""
    import re
    if s is None:
        return None
    s = str(s).replace("，", ",").replace("￥", "").replace("¥", "")
    s = s.replace("元", "").replace(" ", "").strip()
    m = re.search(r"-?\d[\d,]*(?:\.\d+)?", s)
    if not m:
        return None
    return float(m.group(0).replace(",", ""))


if __name__ == "__main__":
    for v in [198.33, 198, 198.3, 0.05, 100000001, 1001.05, 39.8, 11.8, 0.01, 246.93]:
        print(v, "->", rmb_upper(v), "| fmt:", fmt_money(v))
