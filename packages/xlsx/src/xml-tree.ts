const MAX_XML_LENGTH = 32 * 1024 * 1024;
const MAX_XML_NODES = 500_000;
const MAX_XML_DEPTH = 128;

const decodeEntities = (value: string): string => value.replace(/&(?:#(x[0-9a-f]+|\d+)|amp|lt|gt|quot|apos);/gi, (entity, numeric: string | undefined) => {
  if (numeric) {
    const codePoint = Number.parseInt(numeric.startsWith("x") ? numeric.slice(1) : numeric, numeric.startsWith("x") ? 16 : 10);
    return Number.isSafeInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : "�";
  }
  return ({ "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": "\"", "&apos;": "'" } as Record<string, string>)[entity.toLowerCase()] ?? entity;
});

class XmlTextNode {
  readonly nodeType = 3;

  constructor(readonly text: string) {}

  get textContent(): string {
    return this.text;
  }

  toString(): string {
    return this.text;
  }
}

class XmlElementNode {
  readonly nodeType = 1;
  readonly childNodes: Array<XmlElementNode | XmlTextNode> = [];

  constructor(readonly tagName: string, private readonly attributes: Map<string, string>) {}

  get textContent(): string {
    return this.childNodes.map((child) => child.textContent).join("");
  }

  getAttribute(name: string): string | null {
    return this.attributes.get(name) ?? null;
  }

  hasAttribute(name: string): boolean {
    return this.attributes.has(name);
  }

  toString(): string {
    return `<${this.tagName}>${this.textContent}</${this.tagName}>`;
  }
}

const findTagEnd = (content: string, start: number): number => {
  let quote = "";
  for (let index = start; index < content.length; index += 1) {
    const character = content[index]!;
    if (quote) {
      if (character === quote) quote = "";
    } else if (character === "\"" || character === "'") {
      quote = character;
    } else if (character === ">") {
      return index;
    }
  }
  return -1;
};

const parseOpeningTag = (source: string): { name: string; attributes: Map<string, string>; selfClosing: boolean } => {
  const selfClosing = /\/\s*$/.test(source);
  const normalized = source.replace(/\/\s*$/, "").trim();
  const nameMatch = /^([^\s/>]+)/.exec(normalized);
  if (!nameMatch) throw new Error("Malformed XML opening tag.");
  const name = nameMatch[1]!;
  const attributes = new Map<string, string>();
  const remainder = normalized.slice(name.length);
  const attributePattern = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let consumed = "";
  for (const match of remainder.matchAll(attributePattern)) {
    attributes.set(match[1]!, decodeEntities(match[3] ?? match[4] ?? ""));
    consumed += match[0];
  }
  if (remainder.replace(attributePattern, "").trim()) throw new Error("Malformed XML attributes.");
  return { name, attributes, selfClosing };
};

export const createXmlDocument = (content: string): Document => {
  if (content.length > MAX_XML_LENGTH || /<!DOCTYPE/i.test(content)) {
    throw new Error("Unsafe or oversized XML document.");
  }
  const stack: XmlElementNode[] = [];
  let root: XmlElementNode | undefined;
  let nodeCount = 0;
  let offset = 0;
  while (offset < content.length) {
    const opening = content.indexOf("<", offset);
    if (opening < 0) {
      const text = decodeEntities(content.slice(offset));
      if (text && stack.length) stack.at(-1)!.childNodes.push(new XmlTextNode(text));
      break;
    }
    if (opening > offset && stack.length) {
      stack.at(-1)!.childNodes.push(new XmlTextNode(decodeEntities(content.slice(offset, opening))));
    }
    if (content.startsWith("<!--", opening)) {
      const end = content.indexOf("-->", opening + 4);
      if (end < 0) throw new Error("Unclosed XML comment.");
      offset = end + 3;
      continue;
    }
    if (content.startsWith("<?", opening)) {
      const end = content.indexOf("?>", opening + 2);
      if (end < 0) throw new Error("Unclosed XML declaration.");
      offset = end + 2;
      continue;
    }
    if (content.startsWith("<![CDATA[", opening)) {
      const end = content.indexOf("]]>", opening + 9);
      if (end < 0 || !stack.length) throw new Error("Malformed XML CDATA.");
      stack.at(-1)!.childNodes.push(new XmlTextNode(content.slice(opening + 9, end)));
      offset = end + 3;
      continue;
    }
    const end = findTagEnd(content, opening + 1);
    if (end < 0) throw new Error("Unclosed XML tag.");
    const source = content.slice(opening + 1, end).trim();
    if (source.startsWith("/")) {
      const closingName = source.slice(1).trim();
      const current = stack.pop();
      if (!current || current.tagName !== closingName) throw new Error("Mismatched XML closing tag.");
    } else if (source.startsWith("!")) {
      throw new Error("Unsupported XML declaration.");
    } else {
      const parsed = parseOpeningTag(source);
      const element = new XmlElementNode(parsed.name, parsed.attributes);
      nodeCount += 1;
      if (nodeCount > MAX_XML_NODES || stack.length >= MAX_XML_DEPTH) throw new Error("XML complexity limit exceeded.");
      if (stack.length) stack.at(-1)!.childNodes.push(element);
      else if (root) throw new Error("Multiple XML roots.");
      else root = element;
      if (!parsed.selfClosing) stack.push(element);
    }
    offset = end + 1;
  }
  if (!root || stack.length) throw new Error("Malformed XML document.");
  return { documentElement: root } as unknown as Document;
};
