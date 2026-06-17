from __future__ import annotations

from typing import Any

from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement


def set_run_typeface(run: Any, typeface: str) -> None:
    run.font.name = typeface
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        element = r_pr.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            r_pr.append(element)
        element.set("typeface", typeface)
