// GENERATED FILE. Do not edit by hand.
//
// packages/editor-api/scripts/generate-conformance.mjs writes this file from
// compat/manifest.json (the selection and the conformance exceptions) and
// compat/reference/workbook.inventory.json (the reference types). Run
// `bun run compat:generate` to write it again.
//
// Each `_assert_*` alias compares one selected Office.js member with the real public type
// of the package. IsExact checks both directions, so a narrower or wider type fails, not
// only an unrelated one. A reference type names the package export of the same name, so a
// missing export fails as a real compiler error. Each `_write_*` function checks that the
// reference write type assigns to the real setter.

import type { Expect, IsExact } from '../type-assert.ts';
import type * as XlsxEditor from '../../src/index.ts';

// run
// Not asserted: manifest.json conformanceExceptions "run".
// The package has no global host to run against, so it exports no top-level run function.
// XlsxEditorRuntime#run takes the same batch callback.

// RequestContext

// RequestContext#workbook
type Ref_RequestContext_workbook = { readonly workbook: XlsxEditor.Workbook };
type Auth_RequestContext_workbook = Pick<XlsxEditor.RequestContext, 'workbook'>;
type _check_RequestContext_workbook = IsExact<Ref_RequestContext_workbook, Auth_RequestContext_workbook>;
type _assert_RequestContext_workbook = Expect<_check_RequestContext_workbook>;

// OfficeExtension.ClientRequestContext

// OfficeExtension.ClientRequestContext#load
type Ref_ClientRequestContext_load = { (object: XlsxEditor.ClientObject, option?: XlsxEditor.LoadOption | string | string[] | undefined): void };
type Auth_ClientRequestContext_load = XlsxEditor.ClientRequestContext['load'];
type _check_ClientRequestContext_load = IsExact<Ref_ClientRequestContext_load, Auth_ClientRequestContext_load>;
type _assert_ClientRequestContext_load = Expect<_check_ClientRequestContext_load>;

// OfficeExtension.ClientRequestContext#sync
type Ref_ClientRequestContext_sync = { <T>(passThroughValue?: T | undefined): Promise<T> };
type Auth_ClientRequestContext_sync = XlsxEditor.ClientRequestContext['sync'];
type _check_ClientRequestContext_sync = IsExact<Ref_ClientRequestContext_sync, Auth_ClientRequestContext_sync>;
type _assert_ClientRequestContext_sync = Expect<_check_ClientRequestContext_sync>;

// OfficeExtension.ClientObject

// OfficeExtension.ClientObject#context
// Not asserted: manifest.json conformanceExceptions "OfficeExtension.ClientObject#context".
// The package declares context read-only and types it as RequestContext. The runtime sets the
// context of a proxy, and a write from a consumer would detach the proxy from its batch. The
// package has one context class, RequestContext; ClientRequestContext is a type alias of it.

// OfficeExtension.ClientObject#isNullObject
// Not asserted: manifest.json conformanceExceptions "OfficeExtension.ClientObject#isNullObject".
// The package declares isNullObject read-only. The runtime sets it after a sync, and a write from a
// consumer would make a null object look real.

// Workbook

// Workbook#worksheets
type Ref_Workbook_worksheets = { readonly worksheets: XlsxEditor.WorksheetCollection };
type Auth_Workbook_worksheets = Pick<XlsxEditor.Workbook, 'worksheets'>;
type _check_Workbook_worksheets = IsExact<Ref_Workbook_worksheets, Auth_Workbook_worksheets>;
type _assert_Workbook_worksheets = Expect<_check_Workbook_worksheets>;

// WorksheetCollection

// WorksheetCollection#items
type Ref_WorksheetCollection_items = { readonly items: XlsxEditor.Worksheet[] };
type Auth_WorksheetCollection_items = Pick<XlsxEditor.WorksheetCollection, 'items'>;
type _check_WorksheetCollection_items = IsExact<Ref_WorksheetCollection_items, Auth_WorksheetCollection_items>;
type _assert_WorksheetCollection_items = Expect<_check_WorksheetCollection_items>;

// WorksheetCollection#getItem
type Ref_WorksheetCollection_getItem = { (key: string): XlsxEditor.Worksheet };
type Auth_WorksheetCollection_getItem = XlsxEditor.WorksheetCollection['getItem'];
type _check_WorksheetCollection_getItem = IsExact<Ref_WorksheetCollection_getItem, Auth_WorksheetCollection_getItem>;
type _assert_WorksheetCollection_getItem = Expect<_check_WorksheetCollection_getItem>;

// WorksheetCollection#getItemOrNullObject
type Ref_WorksheetCollection_getItemOrNullObject = { (key: string): XlsxEditor.Worksheet };
type Auth_WorksheetCollection_getItemOrNullObject = XlsxEditor.WorksheetCollection['getItemOrNullObject'];
type _check_WorksheetCollection_getItemOrNullObject = IsExact<Ref_WorksheetCollection_getItemOrNullObject, Auth_WorksheetCollection_getItemOrNullObject>;
type _assert_WorksheetCollection_getItemOrNullObject = Expect<_check_WorksheetCollection_getItemOrNullObject>;

// WorksheetCollection#getActiveWorksheet
type Ref_WorksheetCollection_getActiveWorksheet = { (): XlsxEditor.Worksheet };
type Auth_WorksheetCollection_getActiveWorksheet = XlsxEditor.WorksheetCollection['getActiveWorksheet'];
type _check_WorksheetCollection_getActiveWorksheet = IsExact<Ref_WorksheetCollection_getActiveWorksheet, Auth_WorksheetCollection_getActiveWorksheet>;
type _assert_WorksheetCollection_getActiveWorksheet = Expect<_check_WorksheetCollection_getActiveWorksheet>;

// Worksheet

// Worksheet#name
// Not asserted: manifest.json conformanceExceptions "Worksheet#name".
// The package declares name read-only: a worksheet rename is outside the supported subset. A write
// fails to compile instead of failing at run time.

// Worksheet#getRange
type Ref_Worksheet_getRange = { (address?: string | undefined): XlsxEditor.Range };
type Auth_Worksheet_getRange = XlsxEditor.Worksheet['getRange'];
type _check_Worksheet_getRange = IsExact<Ref_Worksheet_getRange, Auth_Worksheet_getRange>;
type _assert_Worksheet_getRange = Expect<_check_Worksheet_getRange>;

// Range

// Range#address
type Ref_Range_address = { readonly address: string };
type Auth_Range_address = Pick<XlsxEditor.Range, 'address'>;
type _check_Range_address = IsExact<Ref_Range_address, Auth_Range_address>;
type _assert_Range_address = Expect<_check_Range_address>;

// Range#values
type Ref_Range_values = { values: any[][] };
type Auth_Range_values = Pick<XlsxEditor.Range, 'values'>;
type _check_Range_values = IsExact<Ref_Range_values, Auth_Range_values>;
type _assert_Range_values = Expect<_check_Range_values>;
export function _write_Range_values(target: XlsxEditor.Range, value: any[][]): void {
  target['values'] = value;
}

// Range#numberFormat
// Not asserted: manifest.json conformanceExceptions "Range#numberFormat".
// The package declares numberFormat read-only: a number format write is outside the supported
// subset. A write fails to compile instead of failing at run time.
