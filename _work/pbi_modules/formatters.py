import re


class DAXFormatter:
    """Форматирование DAX-формул в стиле DAX Formatter."""

    @staticmethod
    def format_column(name: str, expression) -> str:
        return DAXFormatter._format_with_name(name, expression)

    @staticmethod
    def format_measure(name: str, expression) -> str:
        return DAXFormatter._format_with_name(name, expression)

    @staticmethod
    def format_table(name: str, expression) -> str:
        return DAXFormatter._format_with_name(name, expression)

    @staticmethod
    def _normalize_expression(expression) -> str:
        if not expression:
            return ""
        if isinstance(expression, list):
            expression = "".join(str(x) for x in expression if x)
        return str(expression).strip()

    @staticmethod
    def _format_with_name(name: str, expression) -> str:
        expr = DAXFormatter._normalize_expression(expression)
        if not expr:
            return "—"
        tokens = DAXFormatter._split_var_return(expr)
        merged = DAXFormatter._merge_tokens(tokens)
        result_lines = [f"{name} ="]
        for token in merged:
            token_stripped = token.strip()
            token_lower = token_stripped.lower()
            if token_lower.startswith('var'):
                var_part = re.sub(r'^var\s+', '', token_stripped, flags=re.IGNORECASE).strip()
                eq_idx = var_part.find("=")
                if eq_idx > 0:
                    var_name = var_part[:eq_idx].strip()
                    var_expr = var_part[eq_idx + 1:].strip()
                    result_lines.append(f"VAR {var_name} =")
                    formatted = DAXFormatter._format_expression(var_expr)
                    for eline in formatted.split('\n'):
                        if eline.strip():
                            result_lines.append("    " + eline)
            elif token_lower.startswith('return'):
                result_lines.append("RETURN")
                return_expr = re.sub(r'^return\s*', '', token_stripped, flags=re.IGNORECASE).strip()
                if return_expr:
                    formatted = DAXFormatter._format_expression(return_expr)
                    for eline in formatted.split('\n'):
                        if eline.strip():
                            result_lines.append("    " + eline)
            else:
                formatted = DAXFormatter._format_expression(token_stripped)
                for eline in formatted.split('\n'):
                    if eline.strip():
                        result_lines.append("    " + eline)
        return '\n'.join(result_lines)

    @staticmethod
    def _split_var_return(expr: str) -> list:
        expr = re.sub(r'\)(var\s)', r') \1', expr, flags=re.IGNORECASE)
        expr = re.sub(r'\)(return\b)', r') \1', expr, flags=re.IGNORECASE)
        expr = re.sub(r'\b(return)(\w)', r'\1 \2', expr, flags=re.IGNORECASE)
        tokens = re.split(r'(\bvar\s+|\breturn\b)', expr, flags=re.IGNORECASE)
        return [t for t in tokens if t.strip()]

    @staticmethod
    def _merge_tokens(tokens: list) -> list:
        merged = []
        i = 0
        while i < len(tokens):
            t = tokens[i].strip().lower()
            if t.startswith('var') or t == 'return':
                combined = tokens[i]
                if i + 1 < len(tokens):
                    combined += tokens[i + 1]
                    i += 1
                merged.append(combined)
            else:
                merged.append(tokens[i])
            i += 1
        return merged

    @staticmethod
    def _format_expression(expr: str) -> str:
        if not expr:
            return ""
        expr = re.sub(r'[\r\n]+', ' ', expr)
        expr = re.sub(r'  +', ' ', expr).strip()
        expr = DAXFormatter._add_spaces_operators(expr)
        expr = re.sub(r'([A-Za-zа-яА-ЯёЁ_]\w*)\s*\(', DAXFormatter._add_space_before_paren, expr)
        expr = re.sub(r'(\w)\s+\(\s*\)', r'\1()', expr)
        expr = DAXFormatter._format_parens_recursive(expr, 0)
        expr = DAXFormatter._apply_operator_spacing(expr)
        return expr

    @staticmethod
    def _add_spaces_operators(expr: str) -> str:
        comments = []
        def save_comment(m):
            comments.append(m.group(0))
            return f"__CMT{len(comments)-1}__"
        expr = re.sub(r'--[^\n]*', save_comment, expr)
        strings = []
        def save_sq(m):
            strings.append(m.group(0))
            return f"__SQ{len(strings)-1}__"
        def save_dq(m):
            strings.append(m.group(0))
            return f"__DQ{len(strings)-1}__"
        expr = re.sub(r"'[^']*'", save_sq, expr)
        expr = re.sub(r'"(?:[^"\\]|\\.)*"', save_dq, expr)
        expr = re.sub(r'\s*<>\s*', ' <> ', expr)
        expr = re.sub(r'\s*<=\s*', ' <= ', expr)
        expr = re.sub(r'\s*>=\s*', ' >= ', expr)
        expr = re.sub(r'(?<![<>=!&|])\s*=\s*(?![<>=!&|])', ' = ', expr)
        expr = re.sub(r'(?<!&)&(?!&)', ' & ', expr)
        expr = re.sub(r'\|\|', '||', expr)
        expr = re.sub(r'&&', '&&', expr)
        expr = re.sub(r'(?<!\+)\+(?!\+)', ' + ', expr)
        expr = re.sub(r'(?<!-)-(?![->])', ' - ', expr)
        expr = re.sub(r'\s*<\s*(?!=)', ' < ', expr)
        expr = re.sub(r'\s*>\s*(?!=)', ' > ', expr)
        expr = re.sub(r'  +', ' ', expr)
        for i, s in enumerate(strings):
            expr = expr.replace(f"__SQ{i}__", s)
            expr = expr.replace(f"__DQ{i}__", s)
        for i, c in enumerate(comments):
            expr = expr.replace(f"__CMT{i}__", c)
        return expr.strip()

    @staticmethod
    def _format_parens_recursive(expr: str, depth: int) -> str:
        result = []
        i = 0
        while i < len(expr):
            if expr[i] == '(':
                close_idx = DAXFormatter._find_matching_paren(expr, i)
                if close_idx != -1:
                    inner = expr[i + 1:close_idx].strip()
                    if inner:
                        args = DAXFormatter._split_top_level(inner, ',')
                        if len(args) > 1:
                            formatted_args = []
                            for arg in args:
                                arg = arg.strip()
                                if arg:
                                    formatted = DAXFormatter._format_parens_recursive(arg, depth + 1)
                                    formatted = DAXFormatter._apply_operator_spacing(formatted)
                                    formatted_args.append(formatted)
                            indent = "    " * (depth + 1)
                            result.append("(" + formatted_args[0] + ",")
                            for fa in formatted_args[1:]:
                                result.append("\n" + indent + fa + ",")
                            result[-1] = result[-1].rstrip(",") + ")"
                        else:
                            formatted = DAXFormatter._format_parens_recursive(inner, depth + 1)
                            formatted = DAXFormatter._apply_operator_spacing(formatted)
                            result.append("(" + formatted + ")")
                    else:
                        result.append("()")
                    i = close_idx + 1
                else:
                    result.append(expr[i])
                    i += 1
            else:
                result.append(expr[i])
                i += 1
        return ''.join(result)

    @staticmethod
    def _apply_operator_spacing(expr: str) -> str:
        comments = []
        def save_comment(m):
            comments.append(m.group(0))
            return f"__CMT{len(comments)-1}__"
        expr = re.sub(r'--[^\n]*', save_comment, expr)
        strings = []
        def save_sq(m):
            strings.append(m.group(0))
            return f"__SQ{len(strings)-1}__"
        def save_dq(m):
            strings.append(m.group(0))
            return f"__DQ{len(strings)-1}__"
        expr = re.sub(r"'[^']*'", save_sq, expr)
        expr = re.sub(r'"(?:[^"\\]|\\.)*"', save_dq, expr)
        expr = re.sub(r'\s*<>\s*', ' <> ', expr)
        expr = re.sub(r'\s*<=\s*', ' <= ', expr)
        expr = re.sub(r'\s*>=\s*', ' >= ', expr)
        expr = re.sub(r'(?<![<>=!&|])\s*=\s*(?![<>=!&|])', ' = ', expr)
        expr = re.sub(r'\)(\s*)(?=[A-Za-zа-яА-ЯёЁ0-9_\[\(])', ') ', expr)
        expr = re.sub(r'(?<!&)&(?!&)', ' & ', expr)
        expr = re.sub(r'\|\|', '||', expr)
        expr = re.sub(r'&&', '&&', expr)
        expr = re.sub(r'(?<!\+)\+(?!\+)', ' + ', expr)
        expr = re.sub(r'(?<!-)-(?![->])', ' - ', expr)
        expr = re.sub(r'\s*<\s*(?!=)', ' < ', expr)
        expr = re.sub(r'\s*>\s*(?!=)', ' > ', expr)
        expr = re.sub(r'  +', ' ', expr)
        for i, s in enumerate(strings):
            expr = expr.replace(f"__SQ{i}__", s)
            expr = expr.replace(f"__DQ{i}__", s)
        for i, c in enumerate(comments):
            expr = expr.replace(f"__CMT{i}__", c)
        return expr.strip()

    @staticmethod
    def _split_top_level(expr: str, delimiter: str) -> list:
        parts = []
        depth = 0
        in_sq = False
        in_dq = False
        current = ''
        for i, ch in enumerate(expr):
            if ch == "'" and not in_dq:
                in_sq = not in_sq
                current += ch
            elif ch == '"' and not in_sq:
                if i > 0 and expr[i - 1] == '\\':
                    current += ch
                else:
                    in_dq = not in_dq
                    current += ch
            elif not in_sq and not in_dq:
                if ch in '([{':
                    depth += 1
                    current += ch
                elif ch in ')]}':
                    depth -= 1
                    current += ch
                elif ch == delimiter and depth == 0:
                    parts.append(current)
                    current = ''
                else:
                    current += ch
            else:
                current += ch
        if current.strip():
            parts.append(current)
        return parts

    @staticmethod
    def _find_matching_paren(expr: str, open_idx: int) -> int:
        depth = 1
        in_sq = False
        in_dq = False
        i = open_idx + 1
        while i < len(expr) and depth > 0:
            ch = expr[i]
            if ch == "'" and not in_dq:
                in_sq = not in_sq
            elif ch == '"' and not in_sq:
                in_dq = not in_dq
            elif not in_sq and not in_dq:
                if ch == '(':
                    depth += 1
                elif ch == ')':
                    depth -= 1
                    if depth == 0:
                        return i
            i += 1
        return -1

    @staticmethod
    def _add_space_before_paren(m):
        name = m.group(1)
        if name.startswith('#') or name.isdigit():
            return m.group(0)
        return name + ' ('

    @staticmethod
    def safe_format(formula):
        try:
            return DAXFormatter.format_dax_formula(formula)
        except Exception:
            return str(formula) if formula else "—"

    @staticmethod
    def format_dax_formula(formula):
        if not formula:
            return "—"
        if isinstance(formula, list):
            formula = "".join(str(x) for x in formula if x)
        formula = str(formula).strip()
        if not formula or formula == "Нет формулы":
            return formula if formula else "—"
        formula = re.sub(r'[\r\n\t]+', ' ', formula)
        formula = re.sub(r'  +', ' ', formula)
        measure_name = ""
        eq_match = re.match(r'^((?!var\b|return\b)[a-zA-Zа-яА-ЯёЁ0-9_\s!?]+?)\s*=\s*(.*)', formula, re.DOTALL | re.IGNORECASE)
        if eq_match:
            prefix = eq_match.group(1).strip()
            if len(prefix) < 50:
                measure_name = prefix + " = "
                formula = eq_match.group(2).strip()
        tokens = DAXFormatter._split_var_return(formula)
        merged = DAXFormatter._merge_tokens(tokens)
        result_lines = []
        if measure_name.strip():
            result_lines.append(measure_name.strip())
        for token in merged:
            token_stripped = token.strip()
            token_lower = token_stripped.lower()
            if token_lower.startswith('var'):
                var_part = re.sub(r'^var\s+', '', token_stripped, flags=re.IGNORECASE).strip()
                eq_idx = var_part.find("=")
                if eq_idx > 0:
                    var_name = var_part[:eq_idx].strip()
                    var_expr = var_part[eq_idx + 1:].strip()
                    result_lines.append(f"VAR {var_name} =")
                    formatted = DAXFormatter._format_expression(var_expr)
                    for eline in formatted.split('\n'):
                        if eline.strip():
                            result_lines.append("    " + eline)
            elif token_lower.startswith('return'):
                result_lines.append("RETURN")
                return_expr = re.sub(r'^return\s*', '', token_stripped, flags=re.IGNORECASE).strip()
                if return_expr:
                    formatted = DAXFormatter._format_expression(return_expr)
                    for eline in formatted.split('\n'):
                        if eline.strip():
                            result_lines.append("    " + eline)
            else:
                formatted = DAXFormatter._format_expression(token_stripped)
                for eline in formatted.split('\n'):
                    if eline.strip():
                        result_lines.append("    " + eline)
        return '\n'.join(result_lines)


class SQLFormatter:
    """Форматирование SQL внутри M-кода Power Query."""

    @staticmethod
    def format_sql_in_m(m_code: str) -> str:
        if not m_code or not isinstance(m_code, str):
            return m_code
        code = m_code.replace('#(lf)', '\n')
        code = code.replace('#(cr)', '')
        code = code.replace('#(tab)', '    ')
        code = SQLFormatter._format_embedded_sql(code)
        code = SQLFormatter._format_m_steps(code)
        return code

    @staticmethod
    def _format_embedded_sql(code: str) -> str:
        pattern = r'(Query\s*=\s*")((?:[^"\\]|\\.)*?)(\])'
        def replacer(m):
            prefix = m.group(1)
            sql = m.group(2)
            suffix = m.group(3)
            formatted = SQLFormatter._beautify_sql(sql)
            return prefix + formatted + suffix
        return re.sub(pattern, replacer, code, flags=re.DOTALL)

    @staticmethod
    def _beautify_sql(sql: str) -> str:
        if not sql or not sql.strip():
            return sql
        sql = re.sub(r'\s+', ' ', sql)
        clauses = [
            'SELECT', 'FROM', 'WHERE',
            'LEFT JOIN', 'RIGHT JOIN', 'INNER JOIN', 'CROSS JOIN', 'FULL JOIN', 'JOIN',
            'ON', 'GROUP BY', 'HAVING', 'ORDER BY', 'LIMIT', 'OFFSET',
            'UNION ALL', 'UNION',
        ]
        result = sql
        for clause in clauses:
            regex = r'\s+' + re.escape(clause) + r'\s+'
            if re.search(regex, result, re.IGNORECASE):
                result = re.sub(regex, f'\n{clause} ', result, flags=re.IGNORECASE)
        result = re.sub(r'^\s*', '', result)
        select_from = re.search(r'(SELECT\s+)(.*?)(\s+FROM\b)', result, re.DOTALL | re.IGNORECASE)
        if select_from:
            select_kw = select_from.group(1)
            fields_str = select_from.group(2)
            from_kw = select_from.group(3)
            fields = [f.strip() for f in fields_str.split(',') if f.strip()]
            if len(fields) > 1:
                formatted_fields = ',\n    '.join(fields)
                result = result[:select_from.start()] + select_kw + formatted_fields + from_kw + result[select_from.end():]
        for kw in ['AS', 'ON', 'AND', 'OR', 'NOT', 'IN', 'LIKE', 'BETWEEN', 'IS', 'NULL',
                     'LEFT', 'RIGHT', 'INNER', 'OUTER', 'CROSS', 'FULL', 'JOIN',
                     'SELECT', 'FROM', 'WHERE', 'GROUP BY', 'HAVING', 'ORDER BY']:
            pattern = r'\b' + re.escape(kw) + r'\b'
            result = re.sub(pattern, kw, result, flags=re.IGNORECASE)
        return result

    @staticmethod
    def _format_m_steps(code: str) -> str:
        code = re.sub(r',\s*(#"[\w\s]+")', r',\n    \1', code)
        code = re.sub(r'([\)\}\]])\s*in\s*', r'\1\nin\n    ', code)
        code = re.sub(r'\s+in\s+', '\nin\n    ', code)
        code = re.sub(r'\blet\s+', 'let\n    ', code)
        return code


class SourceFormatter:
    """Форматирование источников: DAX или M+SQL."""

    @staticmethod
    def format_source(src_type: str, expression: str) -> str:
        if not expression:
            return expression
        if src_type == "calculated":
            return DAXFormatter.safe_format(expression)
        elif src_type == "m":
            return SQLFormatter.format_sql_in_m(expression)
        else:
            return expression


# ============================================================
# 2. Dependency Analyzer
# ============================================================

