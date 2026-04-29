# Future implementation idea (TODO): VBA Extension Plan

This document outlines a staged plan for adding VBA-only features to pybasil without breaking existing VBScript 5.8 behavior.

## Goal

The primary compatibility target remains VBScript 5.8. VBA-compatible features are desirable when they can be added without changing the parsing or runtime behavior of valid VBScript programs.

In practice this means:

- VBScript behavior stays authoritative for overlapping syntax.
- VBA-only syntax is acceptable when VBScript would reject it anyway.
- Additions should avoid changing keyword handling, operator behavior, or coercion rules for existing VBScript code.
- When a feature introduces ambiguity, it should be constrained to a specific syntactic context or made opt-in.

## Rule of Thumb

Safe to add first:

- syntax forms that VBScript rejects anyway
- builtins or runtime objects that are purely additive
- declaration metadata that can be parsed without changing execution
- context-limited grammar changes

Riskier to add:

- features that reserve new keywords in existing VBScript contexts
- features that alter parsing of already-valid VBScript
- features that change established VBScript coercion or evaluation rules
- features that depend on host applications such as Excel or Word

## Stage 1: Very Safe Syntax Extensions

### 1. `Case ... To ...`

Examples:

- `Case 1 To 5`

Why it is a good first target:

- valid in VBA
- rejected by VBScript
- isolated to `Select Case`

Main risk:

- parser ambiguity inside `Select Case`

Guardrail:

- only recognize this form within `Case` clauses

### 2. `Case Is ...`

Examples:

- `Case Is > 10`
- `Case Is <= x`

Why it is a good first target:

- valid in VBA
- rejected by VBScript
- isolated to `Select Case`

Main risk:

- parser ambiguity inside `Select Case`

Guardrail:

- only recognize this form within `Case` clauses

### 3. Typed declarations

Examples:

- `Dim x As Integer`
- `Dim name As String`
- `Private count As Long`

Why it is a good first target:

- VBScript rejects typed declarations
- initial implementation can treat type annotations as metadata only

Main risk:

- introducing `As` too broadly can break identifier parsing

Guardrail:

- only parse `As <type>` in declaration contexts
- do not make `As` a general-purpose keyword in expressions

### 4. Typed parameters and return annotations

Examples:

- `Function F(x As Long) As String`
- `Sub S(ByVal n As Integer)`

Why it is a good first target:

- invalid in VBScript
- useful for VBA compatibility
- can be metadata-only at first

Main risk:

- same contextual `As` handling concerns as typed declarations

Guardrail:

- support only in procedure signatures

## Stage 2: Safe Parser-First / Light-Semantic Features

### 5. `Option Compare Binary` / `Option Compare Text`

Why it is useful:

- common VBA directive
- self-contained

Main risk:

- changes string comparison semantics globally within a module

Guardrail:

- default remains VBScript behavior unless the directive is explicitly present
- test carefully around comparisons, `Select Case`, `InStr`, `Replace`, and sorting-like behaviors

### 6. `Public` / `Private` on more declarations

Examples:

- `Public Sub Foo()`
- `Private Function Bar()`
- `Public x As Long`

Why it is useful:

- common VBA structure feature
- mostly declaration metadata initially

Main risk:

- enforcing visibility semantics may require module-level runtime structure

Guardrail:

- parse first
- enforce visibility only when the surrounding model is ready

### 7. `Static` declarations

Examples:

- `Static counter As Long`
- `Static x`

Why it is useful:

- common VBA procedure feature

Main risk:

- changes runtime lifetime and state retention

Guardrail:

- start with parser support or tightly scoped semantics
- do not silently change ordinary local-variable behavior

## Stage 3: Host-Independent Runtime Additions

### 8. `Collection` object

Why it is useful:

- common VBA runtime object
- additive feature
- does not change VBScript parsing

Main risk:

- runtime semantics and iteration behavior

Guardrail:

- model behavior explicitly
- add comparison tests against a real VBA host when possible

### 9. Additional VBA intrinsic functions

Good candidates:

- `Choose`
- `Switch`
- `IIf`

Why they are useful:

- additive builtins
- low parser risk

Main risk:

- some names may need care if they overlap with user expectations or future syntax

Guardrail:

- keep them as builtins only
- add targeted tests around evaluation order and coercion

### 10. Expanded scalar type metadata

Examples:

- `Byte`
- `Single`
- `Currency`
- `Decimal`

Why it is useful:

- complements typed declarations
- can begin as metadata support before full coercion support

Main risk:

- full VBA type semantics are deeper than parser support alone

Guardrail:

- parse first
- stage runtime enforcement later

## Stage 4: Useful but Higher-Risk Features

### 11. Fuller property/default-member VBA behavior

Why it is useful:

- improves object-model compatibility

Main risk:

- call/property ambiguity
- subtle runtime dispatch changes

### 12. `For Each` over richer VBA collections

Why it is useful:

- natural follow-on once `Collection` exists

Main risk:

- runtime protocol design

### 13. `WithEvents`, `Implements`, and more advanced class features

Why they are lower priority:

- meaningful complexity increase
- likely to require a more VBA-shaped runtime model

## Features to Avoid Early

These are poor early targets because they are host-bound or likely to alter existing semantics too much.

### Host-bound features

Examples:

- `Debug.Print`
- `DoEvents`
- `Application`
- `Worksheet`
- `Range`
- UserForms
- event procedures tied to Office hosts

Reason:

- they depend on a VBA host environment, not just language syntax

### Broad keyword additions

Examples:

- `As` outside declaration contexts
- `Optional`
- `ParamArray`
- `Friend`

Reason:

- they can easily break currently valid VBScript parsing

### Global semantic changes

Examples:

- changing default comparison behavior globally
- replacing VBScript coercion rules with VBA rules in shared syntax

Reason:

- VBScript remains the baseline for overlapping language features

## Recommended Implementation Order

### Phase A

1. `Case ... To ...`
2. `Case Is ...`
3. non-regression tests proving unchanged VBScript parsing nearby

### Phase B

4. typed `Dim`
5. typed parameters
6. typed return annotations
7. metadata-only handling first

### Phase C

8. `Option Compare Binary/Text`
9. carefully scoped semantic changes only when the directive is present

### Phase D

10. `Collection`
11. additive builtins such as `Choose`, `Switch`, `IIf`

## Guardrails for Every Feature

For each VBA-only addition:

1. add VBScript non-regression tests in the same syntax neighborhood
2. add dedicated VBA-extension tests
3. use `cscript.exe` as the oracle for overlapping VBScript behavior
4. use a real VBA host as the oracle for VBA-only behavior when needed
5. prefer parser support first and semantics second
6. constrain new keywords to specific contexts whenever possible

## Best Safe First Batch

If the goal is to add useful VBA support with minimal risk, the best first batch is:

- `Case 1 To 5`
- `Case Is > 10`
- `Dim x As Integer`
- typed parameters and typed return annotations
- `Option Compare Text`
- `Collection`
- `Choose`, `Switch`, `IIf`

This gives meaningful VBA compatibility without dragging in Office-host-specific runtime behavior.
