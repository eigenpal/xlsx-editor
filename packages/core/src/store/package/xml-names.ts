// Taken through pptx-editor/packages/core/src/store/package/xml-names.ts at 1943d2c9 (EigenPal).
// Difference: the Part 3 citations use the subclause numbers of the 5th edition (7.2 to 7.7).
// XML names and namespace bindings, checked the same way on read, on every op, and on save.
//
// Names follow XML 1.0 (Fifth Edition) §2.3 `NameStartChar`/`NameChar`, without the colon, which
// is the NCName of Namespaces in XML 1.0 §3. The binding rules are Namespaces in XML 1.0 §3: the
// `xmlns` prefix is never declared or used on a name, the `xml` prefix means only the XML
// namespace, and neither reserved namespace is bound to another prefix.
//
// Difference: `prefixValueError` also checks the prefixes inside attribute values (the Markup
// Compatibility attributes of ECMA-376 Part 3, 7.2 to 7.7, and `xsi:type`), as
// docx-editor/packages/core/src/store/package/ooxml-shared.ts (`canonicalQNameAttributeValue`)
// does on read and on save. Unlike that file, a `prefix:*` token of a QName list is allowed.
// The canonical fingerprint (`canonicalForm`) still compares these values as text, where that
// file resolves them to URIs (ECMA-376 Part 3, 7.2 to 7.7).

import { NS } from './namespaces.ts';

const START =
  'A-Z_a-z\\u00C0-\\u00D6\\u00D8-\\u00F6\\u00F8-\\u02FF\\u0370-\\u037D\\u037F-\\u1FFF' +
  '\\u200C-\\u200D\\u2070-\\u218F\\u2C00-\\u2FEF\\u3001-\\uD7FF\\uF900-\\uFDCF\\uFDF0-\\uFFFD' +
  '\\u{10000}-\\u{EFFFF}';
const CHAR = `${START}\\-.0-9\\u00B7\\u0300-\\u036F\\u203F-\\u2040`;
const NCNAME = new RegExp(`^[${START}][${CHAR}]*$`, 'u');

/** Whether `value` is an XML NCName: a name with no colon. */
export function isNcName(value: unknown): value is string {
  return typeof value === 'string' && NCNAME.test(value);
}

/**
 * A broken namespace rule: `reserved` for a misuse of the `xml` or `xmlns` prefix or namespace,
 * `namespace` for a binding XML 1.0 cannot express.
 */
export interface NameRuleViolation {
  readonly rule: 'reserved' | 'namespace';
  readonly detail: string;
}

const reserved = (detail: string): NameRuleViolation => ({ rule: 'reserved', detail });
const namespace = (detail: string): NameRuleViolation => ({ rule: 'namespace', detail });

/**
 * Why a namespace declaration is not allowed, or null when it is.
 *
 * `prefix` is '' for the default namespace. An empty `uri` undeclares the default namespace;
 * undeclaring a prefix is XML 1.1 only and is refused.
 */
export function namespaceBindingError(prefix: string, uri: string): NameRuleViolation | null {
  if (prefix === 'xmlns') return reserved('the xmlns prefix is reserved');
  if (uri === NS.xmlns) return reserved('the xmlns namespace cannot be declared');
  if (prefix === 'xml')
    return uri === NS.xml ? null : reserved('the xml prefix is bound to the XML namespace');
  if (uri === NS.xml) return reserved('the XML namespace is bound only to the xml prefix');
  if (prefix !== '' && uri === '') return namespace('a prefix cannot be undeclared');
  return null;
}

/**
 * Why an expanded name with an optional preferred prefix is not allowed, or null when it is.
 *
 * Shared by elements and attributes. An unprefixed attribute named `xmlns` is a namespace
 * declaration, never an attribute; declarations travel in `nsDecls`.
 */
export function qualifiedNameError(
  ns: string,
  local: string,
  prefix: string | undefined,
  isAttribute: boolean
): NameRuleViolation | null {
  if (ns === NS.xmlns) return reserved('names in the xmlns namespace are reserved');
  if (prefix === 'xmlns') return reserved('the xmlns prefix is reserved');
  if (prefix === 'xml' && ns !== NS.xml)
    return reserved('the xml prefix is bound to the XML namespace');
  if (ns === NS.xml && prefix !== undefined && prefix !== '' && prefix !== 'xml')
    return reserved('the XML namespace is bound only to the xml prefix');
  if (isAttribute && ns === '' && local === 'xmlns')
    return reserved('xmlns is a namespace declaration');
  if (ns === '' && prefix !== undefined && prefix !== '')
    return namespace('a prefix needs a namespace');
  return null;
}

const QNAME_LISTS = new Set(['ProcessContent', 'PreserveElements', 'PreserveAttributes']);

/** Why an attribute value that names prefixes is not allowed in `scope`. */
export interface PrefixValueViolation {
  readonly reason: 'unbound-prefix' | 'bad-name';
  readonly detail: string;
}

/**
 * Why an attribute value names a prefix that `scope` does not declare, or a token that is not a
 * name, or null when the value is good or names no prefix. Markup Compatibility needs each
 * prefix of `mc:Ignorable`, `mc:MustUnderstand`, and `mc:Choice/@Requires` (prefix lists), and
 * of `mc:ProcessContent`, `mc:PreserveElements`, and `mc:PreserveAttributes` (lists of
 * qualified names, where `prefix:*` stands for every name), to be declared in scope (ECMA-376
 * Part 3, 7.2 to 7.7; the 5th edition no longer defines the two `mc:Preserve*`
 * attributes). The value of `xsi:type` is one qualified name. `scope` holds the
 * declarations in scope of the element that carries the attribute, its own included.
 */
export function prefixValueError(
  attribute: { readonly ns: string; readonly local: string; readonly value: string },
  ownerNs: string,
  ownerLocal: string,
  scope: ReadonlyMap<string, string>
): PrefixValueViolation | null {
  const { ns, local, value } = attribute;
  const prefixList =
    (ns === NS.mc && (local === 'Ignorable' || local === 'MustUnderstand')) ||
    (ns === '' && local === 'Requires' && ownerNs === NS.mc && ownerLocal === 'Choice');
  const qnameList = ns === NS.mc && QNAME_LISTS.has(local);
  const single = ns === NS.xsi && local === 'type';
  if (!prefixList && !qnameList && !single) return null;
  const tokens = single ? [value.trim()] : value.split(/[ \t\r\n]+/).filter(Boolean);
  for (const token of tokens) {
    // A prefix list holds bare prefixes. A qualified name has an optional prefix.
    const colon = prefixList ? token.length : token.indexOf(':');
    const prefix = colon < 0 ? '' : token.slice(0, colon);
    const name = colon < 0 ? token : token.slice(colon + 1);
    const good =
      (colon < 0 || isNcName(prefix)) &&
      (prefixList || isNcName(name) || (qnameList && colon >= 0 && name === '*'));
    if (!good) return { reason: 'bad-name', detail: `${JSON.stringify(token)} is not a name` };
    if (prefix !== '' && !scope.has(prefix))
      return { reason: 'unbound-prefix', detail: `the prefix ${prefix} is not declared` };
  }
  return null;
}
