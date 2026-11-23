## ADDED Requirements

### Requirement: Bounded package read

The system SHALL read an XLSX package only through the bounded zip, OPC name, and XML readers of `@eigenpal/ooxml-core`, and SHALL refuse a package that exceeds a limit with a typed reason.

#### Scenario: A zip bomb is refused

- **WHEN** a package declares entries whose decompressed size exceeds the total size or ratio cap
- **THEN** the read fails with a typed limit reason and allocates no buffer of the declared size

#### Scenario: A part name escapes the package

- **WHEN** a relationship target or zip entry name contains `..`, a leading `/`, or an encoded separator
- **THEN** the read refuses that name and does not resolve it

### Requirement: Canonical ordered tree per part

The system SHALL represent each XML part as one ordered tree in which supported elements are typed nodes and unsupported elements are generic nodes in the same child sequence.

#### Scenario: Known and unknown siblings load together

- **WHEN** a worksheet holds supported elements beside an unsupported element in `x:extLst`
- **THEN** the canonical tree holds both in source order, without a second preservation model

#### Scenario: A misplaced known element demotes

- **WHEN** a known element appears where the schema does not allow it
- **THEN** the tree keeps it as a generic node and editing of the rest of the part stays available

### Requirement: Sparse cell index

The system SHALL derive cell lookups only from the cells that exist in the file, and SHALL NOT allocate or iterate by a size or range that a file value states.

#### Scenario: A false dimension does not allocate

- **WHEN** a sheet declares `x:dimension ref="A1:XFD1048576"` and holds three cells
- **THEN** loading the sheet does work in proportion to the three cells

#### Scenario: Other file-stated sizes do not allocate

- **WHEN** a sheet holds `x:mergeCell ref="A1:XFD1048576"`, `x:col min="1" max="16384"`, a shared formula with `ref="A1:XFD1048576"`, rows with `spans="1:16384"`, and the shared string table states `count="2147483647"`
- **THEN** loading does work in proportion to the elements that exist, and no allocation or loop bound comes from those values

#### Scenario: Too many elements

- **WHEN** a sheet holds more rows or cells than the element count cap
- **THEN** the read refuses the part with a typed limit reason

### Requirement: Untrusted values reach sinks only sanitized

The system SHALL sanitize file-derived values at the parse boundary: hyperlink targets through `sanitizeHref`, strings through `escapeXml` on save, and file-keyed maps through `Map` or null-prototype records. It SHALL NOT load an external relationship, an external workbook link, or a data connection without an explicit user action.

#### Scenario: A script hyperlink

- **WHEN** a cell hyperlink targets `javascript:alert(1)`
- **THEN** the projection holds no `href` for the hyperlink, and the target is preserved inert in the part

#### Scenario: A hostile sheet name

- **WHEN** a sheet is named `__proto__`
- **THEN** sheet lookups work and no object prototype changes

#### Scenario: An external link

- **WHEN** a workbook has an `x:externalLink` part
- **THEN** opening the workbook makes no network request, and the part is preserved

#### Scenario: Active content stays inert

- **WHEN** a workbook holds `xl/vbaProject.bin`, an `x:ddeLink`, an OLE object, and an embedded control
- **THEN** opening and saving the workbook runs none of them, and each part is in the saved package byte for byte or as a generic node

### Requirement: Atomic typed mutation

The `WorkbookDocumentStore` SHALL accept changes only as typed ops in a transaction, and SHALL publish a new tree revision atomically or refuse the whole transaction with a typed reason.

#### Scenario: An op that leaves a part invalid

- **WHEN** a transaction would break a structural rule of the typed model in a touched part (child order, a required attribute, or an index into a shared part)
- **THEN** the store refuses the transaction with a typed reason, the tree revision does not change, and the check costs time in proportion to the touched content

#### Scenario: A refused op has no effect

- **WHEN** one op of a transaction targets a cell reference outside the grid
- **THEN** the store refuses the transaction and the tree revision does not change

### Requirement: Normalized save with fidelity oracles

The system SHALL serialize each modeled part in schema order from the canonical tree, SHALL copy each non-XML part byte for byte, and SHALL keep the canonical fingerprint and the save-and-reopen semantic digest equal for every part that no edit and no save rule (design D3 and D5) changes.

#### Scenario: Open and save without an edit

- **WHEN** a fixture is opened and saved with no edit
- **THEN** the saved package reopens with the same fingerprint and digest for every part, and its modeled elements validate against the schemas, with content under `mc:Ignorable` namespaces and `x:extLst` checked only for well-formedness, and `t="d"` recorded as a known schema exception

#### Scenario: Unsupported content survives an edit

- **WHEN** a cell next to a preserved generic element is edited and the workbook is saved
- **THEN** the generic element is in the saved part with its content and relative order
