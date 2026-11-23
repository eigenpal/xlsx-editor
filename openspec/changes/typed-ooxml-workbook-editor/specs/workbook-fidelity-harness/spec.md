## ADDED Requirements

### Requirement: Reference PDFs for every rendering feature

Every rendering feature SHALL land with a synthetic fixture built by a script and its reference PDF, and the harness SHALL render each fixture through the editor's own layout and paint.

#### Scenario: A missing reference is a gap

- **WHEN** a fixture has no current reference PDF
- **THEN** the harness reports a gap and the gate does not pass

#### Scenario: A stale reference is a gap

- **WHEN** the hash of a fixture differs from the hash that `fixtures/references.json` records for its reference PDF
- **THEN** the harness reports the reference as stale, and the gate does not pass

### Requirement: Approximations are listed once

The project SHALL name every rendering approximation in one list in code, and a test SHALL keep that list equal to the public known approximations.

#### Scenario: An approximation without documentation

- **WHEN** a change adds an entry to the code list and not to the public known approximations
- **THEN** the test fails and names the entry

### Requirement: Thresholds only tighten

The harness SHALL keep its thresholds and known gaps in one file, copied from PPTX Editor, and SHALL fail when a threshold or a gap becomes looser than at the merge base, when a fixture that has a gap is renamed or removed, when a new gap appears without a new fixture, or when a score improves beyond its tolerance and its recorded gap does not move. Scores SHALL be recorded for each font environment.

#### Scenario: A looser value

- **WHEN** a change makes a fixture score worse than its known gap
- **THEN** the gate fails and names the fixture, the check, and both values

#### Scenario: A better score

- **WHEN** a change makes a fixture score better than its known gap by more than the tolerance of its check
- **THEN** the gate fails until the change records the new value

#### Scenario: A renamed fixture

- **WHEN** a change renames a fixture that has a known gap
- **THEN** the gate fails and names the old fixture

### Requirement: Save check

The project SHALL open saved and edited copies of each fixture in a local check, and SHALL record a fixture as save-verified, with its hash and the serializer version, only when both copies open with no repair prompt and with every sheet. CI SHALL fail when a record does not match the current fixture hash or serializer version.

#### Scenario: A repair prompt

- **WHEN** a saved copy opens with a repair prompt
- **THEN** the fixture is not save-verified and the defect is filed with the `priority:high` label and the matching `area:*` label
