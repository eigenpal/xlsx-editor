// The kind of a workbook main part, from its content type (ECMA-376 Part 2, 7.2.3). ECMA-376
// Part 1, 12.3.23 lists the workbook and the template content types. A package can also use
// their macro-enabled forms: the main part has the same root element (`x:workbook`,
// `CT_Workbook`) and the same `officeDocument` relationship, so the engine reads all four the
// same way. A macro part stays as bytes that are never run.
//
// `store/index.ts` does not export this module: hosts read `snapshot().workbookKind` and
// `snapshot().macroEnabled`.

import { asciiFold } from '@eigenpal/ooxml-core';
import { CONTENT_TYPE } from './namespaces.ts';

const folded = (...types: string[]) => new Set(types.map(asciiFold));

/** The macro-enabled forms of the workbook and the template main part. */
const MACRO_WORKBOOK = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
const MACRO_TEMPLATE = 'application/vnd.ms-excel.template.macroEnabled.main+xml';

/** The main-part content types that the engine opens. Others, such as an add-in, are refused. */
export const WORKBOOK_TYPES: ReadonlySet<string> = folded(
  CONTENT_TYPE.workbook,
  CONTENT_TYPE.template,
  MACRO_WORKBOOK,
  MACRO_TEMPLATE
);
const TEMPLATE_TYPES = folded(CONTENT_TYPE.template, MACRO_TEMPLATE);
const MACRO_TYPES = folded(MACRO_WORKBOOK, MACRO_TEMPLATE);

const has = (set: ReadonlySet<string>, type: string | undefined): boolean =>
  type !== undefined && set.has(asciiFold(type.trim()));

/** Whether a main-part content type is a workbook type that the engine opens. */
export const isWorkbookType = (type: string | undefined): boolean => has(WORKBOOK_TYPES, type);

/** Whether a main-part content type is a template, macro-enabled or not. */
export const isTemplateType = (type: string | undefined): boolean => has(TEMPLATE_TYPES, type);

/** Whether a main-part content type is macro-enabled. */
export const isMacroEnabledType = (type: string | undefined): boolean => has(MACRO_TYPES, type);
