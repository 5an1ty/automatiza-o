"""Conservative product text boundaries shared by extraction and document output."""
import re
import unicodedata


def normalize(value):
    return re.sub(r'\s+', ' ', ''.join(c for c in unicodedata.normalize('NFD', value.lower())
                  if unicodedata.category(c) != 'Mn')).strip()


def commercial_line(text):
    text = normalize(text)
    # Section headings and field labels, rather than words anywhere in a product.
    # A service such as "Transporte rodoviário de carga" remains a valid item.
    if re.match(r'^(?:condic(?:ao|oes) (?:comercia(?:l|is)|de pagamento)|dados bancarios|'
                r'observacoes(?: / condicoes| comerciais)?|informacoes comerciais|'
                r'prazo(?:s)? (?:de |para )?(?:entrega|pagamento)|'
                r'forma de pagamento|valor total(?: do orcamento)?|subtotal)\b', text):
        return True
    if re.match(r'^(?:prazo|pagamento|frete|transporte|banco|agencia|conta(?: corrente)?|'
                r'chave pix|pix|validade(?: da cotacao)?|total|entrega)\s*[:=–—-]', text):
        return True
    if re.match(r'^\d+(?:\s*(?:a|/|-)\s*\d+)*\s*(?:dias?|semanas?|meses?|horas?)\b', text):
        return True
    if text in ('pagamento', 'transporte', 'frete', 'validade', 'total'):
        return True
    return bool(re.match(r'^(?:a vista|a prazo|boleto|fob|cif|'
                         r'pagamento|transferencia bancaria|frete (?:incluso|incluido|por conta)|'
                         r'entrega (?:em|apos|mediante|prevista)|apos (?:aprovacao|recebimento)|'
                         r'pedido minimo|sujeito a (?:aprovacao|disponibilidade))\b', text))


def clean_description(text):
    lines = []
    for line in text.splitlines():
        line = line.strip()
        if not line or commercial_line(line):
            continue
        for separator in re.finditer(r';\s*|,\s*|\s+[-–—]\s+', line):
            if commercial_line(line[separator.end():]):
                line = line[:separator.start()].strip()
                break
        if re.match(r'^(?:fabricante|(?:codigo (?:do )?fabricante)|quantidade|preco total|valor total)\s*:', normalize(line)):
            continue
        lines.append(line)
    return '\n'.join(lines)
