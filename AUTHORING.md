# AUTHORING (Template Content Controls + Baseline Hashes)

This guide is for template authors maintaining **SalesContractTemplate_PoC.docx**.

## 1) Content controls and tags

Template blocks are stored as Word content controls and tagged:

- `TEMPLATE|HD|<Name>|h<hash>`
- `TEMPLATE|TB|<Name>|h<hash>`
- `TEMPLATE|BP|<Name>|h<hash>`

Only blocks with `TEMPLATE|` or `APPROVED|` are treated as “standard” by the add-in.

## 2) Authoring workflow (required)

Whenever you change text inside any TEMPLATE content control:

1) Edit the content
2) Recompute baseline hash (using the same normalization as the add-in)
3) Update the content control **Tag** to end with `|h<newhash>`
4) Save the template

If you skip step (3), Validate will treat the block as changed.

## 3) Word Desktop steps (recommended)

Use Word Desktop for content control property editing:

- Enable Developer tab
- Select content control → Developer → Properties
- Update Title + Tag
- (Optional) Lock: prevent deletion

## 4) Negotiated Clauses section

Two approaches:

- Wrap instructions in a TEMPLATE BP control (so it can validate Green)
- Leave outside any control (so it remains Red)

## 5) Re-baselining after large edits

If you change many template blocks, recompute hashes for all edited controls and update their Tags.
