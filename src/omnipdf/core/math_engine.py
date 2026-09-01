"""
Math Engine: Detects mathematical formulas, LaTeX notation, Greek symbols,
subscripts/superscripts, and generates native Word OMML (<m:oMath>) equations.
"""

import re
import html
from typing import Optional, Tuple, List
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


class MathEngine:
    """Detects, parses, and converts math expressions to native Office Math Markup Language (OMML)."""

    MATH_NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'
    WORD_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

    GREEK_MAP = {
        r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ",
        r"\epsilon": "ε", r"\zeta": "ζ", r"\eta": "η", r"\theta": "θ",
        r"\iota": "ι", r"\kappa": "κ", r"\lambda": "λ", r"\mu": "μ",
        r"\nu": "ν", r"\xi": "ξ", r"\pi": "π", r"\rho": "ρ",
        r"\sigma": "σ", r"\tau": "τ", r"\upsilon": "υ", r"\phi": "φ",
        r"\chi": "χ", r"\psi": "ψ", r"\omega": "ω",
        r"\Gamma": "Γ", r"\Delta": "Δ", r"\Theta": "Θ", r"\Lambda": "Λ",
        r"\Xi": "Ξ", r"\Pi": "Π", r"\Sigma": "Σ", r"\Phi": "Φ",
        r"\Psi": "Ψ", r"\Omega": "Ω",
    }

    SYMBOL_MAP = {
        r"\le": "≤", r"\leq": "≤", r"\ge": "≥", r"\geq": "≥",
        r"\ne": "≠", r"\neq": "≠", r"\approx": "≈", r"\equiv": "≡",
        r"\pm": "±", r"\mp": "∓", r"\times": "×", r"\cdot": "·",
        r"\div": "÷", r"\in": "∈", r"\notin": "∉", r"\subset": "⊂",
        r"\subseteq": "⊆", r"\cup": "∪", r"\cap": "∩", r"\forall": "∀",
        r"\exists": "∃", r"\nabla": "∇", r"\partial": "∂", r"\infty": "∞",
        r"\to": "→", r"\rightarrow": "→", r"\leftarrow": "←",
        r"\Rightarrow": "⇒", r"\Leftarrow": "⇐", r"\Leftrightarrow": "⇔",
        r"\dots": "…", r"\cdots": "⋯", r"\ldots": "…",
    }

    OPERATORS_MAP = {
        r"\sum": ("∑", "sum"),
        r"\int": ("∫", "int"),
        r"\prod": ("∏", "prod"),
        r"\iint": ("∬", "iint"),
        r"\oint": ("∮", "oint"),
    }

    MATH_TRIGGERS = re.compile(
        r"(\\frac|\\sqrt|\\sum|\\int|\\prod|\\mathbf|\\mathcal|\\mathbb|\\text|"
        r"[\u0370-\u03FF\u2100-\u214F\u2200-\u22FF\u2A00-\u2AFF]|"
        r"[a-zA-Z0-9]\^[0-9a-zA-Z\+\-]|"
        r"[a-zA-Z0-9]_[0-9a-zA-Z\+\-]|"
        r"[=<>±×÷·√∑∏∫∮∇∂∞≈≠≡≤≥∈∉∪∩∀∃])"
    )

    DISPLAY_EQ_PATTERN = re.compile(
        r"^\s*(\([0-9\.]+\)|\[[0-9\.]+\]|\([a-zA-Z0-9\.\-]+\))\s*$"
    )

    @classmethod
    def is_math_span(cls, text: str, font_name: str = "") -> bool:
        """Heuristic to check if a text run or span is mathematical."""
        font_lower = font_name.lower()
        if any(k in font_lower for k in ["math", "cmr", "cmsy", "cmmi", "msbm", "stix", "symbol", "euler"]):
            return True
        if cls.MATH_TRIGGERS.search(text):
            # Check length: if it's very long normal prose, ignore unless high symbol density
            symbols = len(cls.MATH_TRIGGERS.findall(text))
            if symbols > 0 and (len(text) < 40 or (symbols / max(1, len(text.split()))) > 0.3):
                return True
        return False

    @classmethod
    def is_display_equation_block(cls, lines: List[str]) -> bool:
        """Check if a block of text represents an isolated display formula."""
        full_text = " ".join(lines).strip()
        if len(full_text) == 0 or len(full_text) > 300:
            return False
        
        # Check if ends with equation number like (1), (2.3), [4]
        has_eq_num = bool(re.search(r"(\([0-9a-zA-Z\.\-]+\)|\[[0-9]+\])\s*$", full_text))
        has_math_tokens = bool(cls.MATH_TRIGGERS.search(full_text))
        has_equals = "=" in full_text or "≈" in full_text or "≤" in full_text or "≥" in full_text

        return (has_math_tokens and has_equals) or (has_eq_num and has_math_tokens)

    @classmethod
    def convert_to_omml(cls, text: str, is_display: bool = False) -> str:
        """Convert a mathematical formula or LaTeX snippet into valid Word OMML XML."""
        cleaned = text.strip()
        if cleaned.startswith("$") and cleaned.endswith("$") and len(cleaned) > 2:
            cleaned = cleaned[1:-1].strip()
        if cleaned.startswith("$$") and cleaned.endswith("$$") and len(cleaned) > 4:
            cleaned = cleaned[2:-2].strip()

        # Step 1: Replace Greek and LaTeX symbols
        for latex_sym, unicode_sym in cls.GREEK_MAP.items():
            cleaned = re.sub(re.escape(latex_sym) + r"(?![a-zA-Z])", unicode_sym, cleaned)
        for latex_sym, unicode_sym in cls.SYMBOL_MAP.items():
            cleaned = re.sub(re.escape(latex_sym) + r"(?![a-zA-Z])", unicode_sym, cleaned)

        # Step 2: Build OMML XML tree
        omml_body = cls._parse_math_tokens_to_omml(cleaned)

        if is_display:
            return f'<m:oMathPara {cls.MATH_NS}><m:oMath>{omml_body}</m:oMath></m:oMathPara>'
        return f'<m:oMath {cls.MATH_NS}>{omml_body}</m:oMath>'

    @classmethod
    def _parse_math_tokens_to_omml(cls, expr: str) -> str:
        """Parse LaTeX and math expressions into OMML components."""
        elements = []
        i = 0
        n = len(expr)

        while i < n:
            # 0. Handle whitespace
            if expr[i].isspace():
                elements.append('<m:r><m:t xml:space="preserve"> </m:t></m:r>')
                while i < n and expr[i].isspace():
                    i += 1
                continue

            # 1. Check for N-ary Operators first: \sum, \int, \prod
            op_match = None
            for op_cmd, (op_char, op_name) in cls.OPERATORS_MAP.items():
                if expr.startswith(op_cmd, i) or expr.startswith(op_char, i):
                    op_len = len(op_cmd) if expr.startswith(op_cmd, i) else len(op_char)
                    op_match = (op_cmd, op_char, op_name, op_len)
                    break
            if op_match:
                op_cmd, op_char, op_name, op_len = op_match
                i += op_len
                sub_val = ""
                sup_val = ""
                # Optional limits
                while i < n and expr[i].isspace():
                    i += 1
                if i < n and expr[i] == "_":
                    i += 1
                    sub_val, i = cls._extract_token_or_group(expr, i)
                while i < n and expr[i].isspace():
                    i += 1
                if i < n and expr[i] == "^":
                    i += 1
                    sup_val, i = cls._extract_token_or_group(expr, i)

                sub_omml = cls._parse_math_tokens_to_omml(sub_val) if sub_val else "<m:r><m:t></m:t></m:r>"
                sup_omml = cls._parse_math_tokens_to_omml(sup_val) if sup_val else "<m:r><m:t></m:t></m:r>"

                elements.append(
                    f'<m:nary><m:naryPr><m:chr m:val="{op_char}"/><m:limLoc m:val="undOvr"/></m:naryPr>'
                    f'<m:sub>{sub_omml}</m:sub><m:sup>{sup_omml}</m:sup><m:e/></m:nary>'
                )
                continue

            # 2. Check for Fractions: \frac{num}{den}
            if expr.startswith(r"\frac", i):
                i += 5
                num, i = cls._extract_braced_group(expr, i)
                den, i = cls._extract_braced_group(expr, i)
                num_omml = cls._parse_math_tokens_to_omml(num)
                den_omml = cls._parse_math_tokens_to_omml(den)
                elements.append(
                    f'<m:f><m:num>{num_omml}</m:num><m:den>{den_omml}</m:den></m:f>'
                )
                continue

            # 3. Check for Square Roots / Radicals: \sqrt{x} or \sqrt[n]{x}
            if expr.startswith(r"\sqrt", i):
                i += 5
                degree = ""
                if i < n and expr[i] == "[":
                    deg_end = expr.find("]", i)
                    if deg_end != -1:
                        degree = expr[i + 1:deg_end]
                        i = deg_end + 1
                base, i = cls._extract_braced_group(expr, i)
                base_omml = cls._parse_math_tokens_to_omml(base)
                if degree:
                    deg_omml = cls._parse_math_tokens_to_omml(degree)
                    elements.append(
                        f'<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg>{deg_omml}</m:deg><m:e>{base_omml}</m:e></m:rad>'
                    )
                else:
                    elements.append(
                        f'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>{base_omml}</m:e></m:rad>'
                    )
                continue

            # 4. Check for Subscript / Superscript combinations: x_i^2 or x^2_i
            subsup_match = re.match(
                r"^([a-zA-Z0-9\(\)\{\}\\]+)(?:_([a-zA-Z0-9\+\-]+|\{[^}]+\})\^([a-zA-Z0-9\+\-]+|\{[^}]+\})|\^([a-zA-Z0-9\+\-]+|\{[^}]+\})_([a-zA-Z0-9\+\-]+|\{[^}]+\}))",
                expr[i:]
            )
            if subsup_match and not subsup_match.group(1).startswith("\\frac") and not subsup_match.group(1).startswith("\\sqrt"):
                base_val = subsup_match.group(1)
                sub_val = subsup_match.group(2) or subsup_match.group(5)
                sup_val = subsup_match.group(3) or subsup_match.group(4)
                
                sub_val = sub_val.strip("{}")
                sup_val = sup_val.strip("{}")
                
                base_omml = cls._parse_math_tokens_to_omml(base_val)
                sub_omml = cls._parse_math_tokens_to_omml(sub_val)
                sup_omml = cls._parse_math_tokens_to_omml(sup_val)
                
                elements.append(
                    f'<m:sSubSup><m:e>{base_omml}</m:e><m:sub>{sub_omml}</m:sub><m:sup>{sup_omml}</m:sup></m:sSubSup>'
                )
                i += len(subsup_match.group(0))
                continue

            # 5. Check for Subscripts: x_i or x_{i+1}
            sub_match = re.match(r"^([a-zA-Z0-9\(\)\\]+)_([a-zA-Z0-9\+\-]+|\{[^}]+\})", expr[i:])
            if sub_match and not sub_match.group(1).startswith("\\frac") and not sub_match.group(1).startswith("\\sqrt"):
                base_val = sub_match.group(1)
                sub_val = sub_match.group(2).strip("{}")
                base_omml = cls._parse_math_tokens_to_omml(base_val)
                sub_omml = cls._parse_math_tokens_to_omml(sub_val)
                elements.append(
                    f'<m:sSub><m:e>{base_omml}</m:e><m:sub>{sub_omml}</m:sub></m:sSub>'
                )
                i += len(sub_match.group(0))
                continue

            # 6. Check for Superscripts: x^2 or x^{n+1}
            sup_match = re.match(r"^([a-zA-Z0-9\(\)\\]+)\^([a-zA-Z0-9\+\-]+|\{[^}]+\})", expr[i:])
            if sup_match and not sup_match.group(1).startswith("\\frac") and not sup_match.group(1).startswith("\\sqrt"):
                base_val = sup_match.group(1)
                sup_val = sup_match.group(2).strip("{}")
                base_omml = cls._parse_math_tokens_to_omml(base_val)
                sup_omml = cls._parse_math_tokens_to_omml(sup_val)
                elements.append(
                    f'<m:sSup><m:e>{base_omml}</m:e><m:sup>{sup_omml}</m:sup></m:sSup>'
                )
                i += len(sup_match.group(0))
                continue

            # 7. Check for Parentheses / Delimiters: ( ... ) or [ ... ]
            if expr[i] in "([{":
                open_char = expr[i]
                close_char = { "(": ")", "[": "]", "{": "}" }[open_char]
                inner_text, i = cls._extract_delimited_group(expr, i, open_char, close_char)
                inner_omml = cls._parse_math_tokens_to_omml(inner_text)
                elements.append(
                    f'<m:d><m:dPr><m:begChr m:val="{open_char}"/><m:endChr m:val="{close_char}"/></m:dPr>'
                    f'<m:e>{inner_omml}</m:e></m:d>'
                )
                continue

            # 8. Single symbol, operator, or token
            char = expr[i]
            # Accumulate word or operator
            if char.isalnum():
                tok = char
                i += 1
                while i < n and expr[i].isalnum():
                    tok += expr[i]
                    i += 1
                escaped = html.escape(tok)
                elements.append(f'<m:r><m:t>{escaped}</m:t></m:r>')
            else:
                escaped = html.escape(char)
                elements.append(f'<m:r><m:t>{escaped}</m:t></m:r>')
                i += 1

        return "".join(elements)

    @classmethod
    def _extract_braced_group(cls, expr: str, start_idx: int) -> Tuple[str, int]:
        """Extract content inside { ... } starting at or after start_idx."""
        while start_idx < len(expr) and expr[start_idx].isspace():
            start_idx += 1
        if start_idx >= len(expr) or expr[start_idx] != "{":
            # Single character fallback
            if start_idx < len(expr):
                return expr[start_idx], start_idx + 1
            return "", start_idx

        depth = 1
        idx = start_idx + 1
        content = []
        while idx < len(expr) and depth > 0:
            if expr[idx] == "{":
                depth += 1
            elif expr[idx] == "}":
                depth -= 1
                if depth == 0:
                    return "".join(content), idx + 1
            content.append(expr[idx])
            idx += 1

        return "".join(content), idx

    @classmethod
    def _extract_token_or_group(cls, expr: str, start_idx: int) -> Tuple[str, int]:
        """Extract a single token or braced group."""
        while start_idx < len(expr) and expr[start_idx].isspace():
            start_idx += 1
        if start_idx >= len(expr):
            return "", start_idx
        if expr[start_idx] == "{":
            return cls._extract_braced_group(expr, start_idx)
        
        # Single token (letters, numbers)
        token_match = re.match(r"^([a-zA-Z0-9\+\-]+)", expr[start_idx:])
        if token_match:
            tok = token_match.group(1)
            return tok, start_idx + len(tok)
        return expr[start_idx], start_idx + 1

    @classmethod
    def _extract_delimited_group(cls, expr: str, start_idx: int, open_c: str, close_c: str) -> Tuple[str, int]:
        """Extract content enclosed by balanced open_c and close_c."""
        depth = 1
        idx = start_idx + 1
        content = []
        while idx < len(expr) and depth > 0:
            if expr[idx] == open_c:
                depth += 1
            elif expr[idx] == close_c:
                depth -= 1
                if depth == 0:
                    return "".join(content), idx + 1
            content.append(expr[idx])
            idx += 1
        return "".join(content), idx

    @classmethod
    def insert_equation_into_paragraph(cls, paragraph, math_text: str, is_display: bool = False):
        """Insert native OMML equation XML into a python-docx Paragraph."""
        try:
            omml_xml = cls.convert_to_omml(math_text, is_display=is_display)
            math_element = parse_xml(omml_xml)
            paragraph._element.append(math_element)
            return True
        except Exception as e:
            # Fallback to plain text run if XML parsing fails
            run = paragraph.add_run(math_text)
            run.font.name = "Cambria Math"
            run.italic = True
            return False
