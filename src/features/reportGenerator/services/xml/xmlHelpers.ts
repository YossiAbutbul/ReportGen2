export function getAttributeByLocalName(
  element: Element,
  localName: string
): string {
  for (const attribute of Array.from(element.attributes)) {
    if (attribute.localName === localName) return attribute.value;
  }

  return "";
}

export function getFirstChildByLocalName(
  parent: ParentNode,
  localName: string
): Element | null {
  return (
    Array.from(parent.childNodes).find(
      (node): node is Element =>
        node.nodeType === Node.ELEMENT_NODE &&
        (node as Element).localName === localName
    ) ?? null
  );
}

export function getDescendantsByLocalName(
  parent: Document | Element,
  localName: string
): Element[] {
  return Array.from(parent.getElementsByTagName("*")).filter(
    (node): node is Element => node.localName === localName
  );
}