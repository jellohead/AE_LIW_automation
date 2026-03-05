from typing import Optional, TYPE_CHECKING
if TYPE_CHECKING:
    from pptx.text.text import Paragraph


def set_safe_indent(p_pr, attr: str, value: Optional[str]):
    if value is not None:
        try:
            p_pr.set(attr, value)
            print('pause here to debug')
        except Exception as e:
            print(f"⚠️ Failed to set {attr} = {value}: {e}")

def format_paragraph_xml(
    para: "Paragraph",
    *,
    level: int = 0,
    left_indent: Optional[str] = None,
    hanging_indent: Optional[str] = None,
    bullet_char: str = u'\u2022' # \u2022 (•)
) -> None:
    p_pr = para._p.get_or_add_pPr()
    p_pr.set('lvl', str(level))

    if left_indent is not None:
        p_pr.set('marL', str(int(left_indent)))
    if hanging_indent is not None:
        p_pr.set('indent', str(int(hanging_indent)))

    # code to change the bullet character if I decide to implement that feature
    # bu_char = p_pr.find(qn('a:buChar'))
    # if bu_char is None:
    #     bu_char = OxmlElement('a:buChar')
    #     p_pr.append(bu_char)
    # bu_char.set('char', bullet_char)
