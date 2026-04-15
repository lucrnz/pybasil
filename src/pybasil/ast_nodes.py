"""AST Node definitions for VBScript parser."""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional, Union
from enum import Enum, auto


class ComparisonOp(Enum):
    EQ = auto()  # =
    NE = auto()  # <>
    LT = auto()  # <
    GT = auto()  # >
    LE = auto()  # <=
    GE = auto()  # >=
    IS = auto()  # Is


class BinaryOp(Enum):
    ADD = auto()  # +
    SUB = auto()  # -
    MUL = auto()  # *
    DIV = auto()  # /
    INTDIV = auto()  # \
    MOD = auto()  # Mod
    POW = auto()  # ^
    CONCAT = auto()  # &
    AND = auto()  # And
    OR = auto()  # Or
    XOR = auto()  # Xor
    EQV = auto()  # Eqv
    IMP = auto()  # Imp


class UnaryOp(Enum):
    NEG = auto()  # -
    POS = auto()  # +
    NOT = auto()  # Not


class ExitType(Enum):
    FOR = auto()  # Exit For
    DO = auto()  # Exit Do
    SUB = auto()  # Exit Sub
    FUNCTION = auto()  # Exit Function


class LoopConditionType(Enum):
    WHILE = auto()  # While condition
    UNTIL = auto()  # Until condition


class ErrorHandlingMode(Enum):
    """Error handling mode for VBScript."""

    DEFAULT = auto()  # Default - raise errors immediately
    RESUME_NEXT = auto()  # On Error Resume Next - continue on error
    GOTO = auto()  # On Error GoTo - jump to label on error


@dataclass
class ASTNode:
    """Base class for all AST nodes."""

    pass


# Literals
@dataclass
class NumberLiteral(ASTNode):
    value: float


@dataclass
class StringLiteral(ASTNode):
    value: str


@dataclass
class BooleanLiteral(ASTNode):
    value: bool


@dataclass
class NothingLiteral(ASTNode):
    pass


@dataclass
class EmptyLiteral(ASTNode):
    pass


@dataclass
class NullLiteral(ASTNode):
    pass


# Expressions
@dataclass
class Identifier(ASTNode):
    name: str
    _lower: str = field(init=False, repr=False, compare=False)

    def __post_init__(self):
        object.__setattr__(self, '_lower', self.name.lower())


@dataclass
class BinaryExpression(ASTNode):
    left: ASTNode
    operator: BinaryOp
    right: ASTNode


@dataclass
class UnaryExpression(ASTNode):
    operator: UnaryOp
    operand: ASTNode


@dataclass
class ComparisonExpression(ASTNode):
    left: ASTNode
    operator: ComparisonOp
    right: ASTNode


@dataclass
class MemberAccess(ASTNode):
    object: ASTNode
    member: str


@dataclass
class FunctionCall(ASTNode):
    name: str
    arguments: List[ASTNode] = field(default_factory=list)


@dataclass
class MethodCall(ASTNode):
    object: ASTNode
    method: str
    arguments: List[ASTNode] = field(default_factory=list)


@dataclass
class NewExpression(ASTNode):
    class_name: str


# Array Access expression
@dataclass
class ArrayAccess(ASTNode):
    """Array element access like arr(i) or matrix(i, j)."""

    name: str
    indices: List[ASTNode] = field(default_factory=list)


# Dim variable declaration (can be simple or array)
@dataclass
class DimVariable(ASTNode):
    """A variable declaration in a Dim statement."""

    name: str
    dimensions: Optional[List[ASTNode]] = (
        None  # None for simple var, [] for dynamic array, [sizes] for fixed
    )


# Statements
@dataclass
class DimStatement(ASTNode):
    variables: List[DimVariable]


@dataclass
class AssignmentStatement(ASTNode):
    variable: str  # Variable name
    indices: Optional[List[ASTNode]] = None  # For array assignment: arr(i) = value
    expression: ASTNode = None


@dataclass
class SetStatement(ASTNode):
    variable: str  # Variable name
    indices: Optional[List[ASTNode]] = None  # For array assignment: Set arr(i) = obj
    expression: ASTNode = None


@dataclass
class PropertyAssignmentStatement(ASTNode):
    """Property assignment like obj.Property = value or obj.Property("key") = value."""

    target: (
        ASTNode  # The expression being assigned to (e.g., MemberAccess or ArrayAccess)
    )
    expression: ASTNode


@dataclass
class CallStatement(ASTNode):
    name: str
    arguments: List[ASTNode] = field(default_factory=list)


@dataclass
class ExpressionStatement(ASTNode):
    expression: ASTNode


# Parameter for procedures
@dataclass
class Parameter(ASTNode):
    name: str
    is_byref: bool = True  # ByRef is default in VBScript


# Procedure Statements
@dataclass
class SubStatement(ASTNode):
    name: str
    parameters: List[Parameter] = field(default_factory=list)
    body: List[ASTNode] = field(default_factory=list)


@dataclass
class FunctionStatement(ASTNode):
    name: str
    parameters: List[Parameter] = field(default_factory=list)
    body: List[ASTNode] = field(default_factory=list)


# Control Flow Statements
@dataclass
class ElseIfClause(ASTNode):
    condition: ASTNode
    body: List[ASTNode]


@dataclass
class ElseClause(ASTNode):
    body: List[ASTNode]


@dataclass
class IfStatement(ASTNode):
    condition: ASTNode
    then_body: List[ASTNode]
    elseif_clauses: List[ElseIfClause] = field(default_factory=list)
    else_clause: Optional[ElseClause] = None


@dataclass
class CaseRange(ASTNode):
    """A Case range expression like Case 1 To 10."""

    low: ASTNode
    high: ASTNode


@dataclass
class CaseComparison(ASTNode):
    """A Case Is comparison like Case Is > 5."""

    operator: ComparisonOp
    expression: ASTNode


@dataclass
class CaseClause(ASTNode):
    """A Case clause in a Select Case statement."""

    values: List[ASTNode]  # List of values, CaseRange, or CaseComparison to match
    body: List[ASTNode]


@dataclass
class CaseElseClause(ASTNode):
    """A Case Else clause in a Select Case statement."""

    body: List[ASTNode]


@dataclass
class SelectCaseStatement(ASTNode):
    """Select Case expression ... Case ... End Select"""

    expression: ASTNode
    case_clauses: List[CaseClause] = field(default_factory=list)
    case_else_clause: Optional[CaseElseClause] = None


@dataclass
class ForStatement(ASTNode):
    variable: str
    start: ASTNode
    end: ASTNode
    step: Optional[ASTNode] = None
    body: List[ASTNode] = field(default_factory=list)


@dataclass
class ForEachStatement(ASTNode):
    """For Each item In collection ... Next"""

    variable: str
    collection: ASTNode
    body: List[ASTNode] = field(default_factory=list)


@dataclass
class ReDimStatement(ASTNode):
    """ReDim [Preserve] name(dims) [, name2(dims2)]*"""

    preserve: bool
    arrays: List[tuple]  # List of (name, dimensions) tuples


@dataclass
class EraseStatement(ASTNode):
    """Erase array1 [, array2]*"""

    arrays: List[str]


@dataclass
class WhileStatement(ASTNode):
    condition: ASTNode
    body: List[ASTNode] = field(default_factory=list)


@dataclass
class LoopCondition(ASTNode):
    condition_type: LoopConditionType
    condition: ASTNode


@dataclass
class DoLoopStatement(ASTNode):
    pre_condition: Optional[LoopCondition] = None
    body: List[ASTNode] = field(default_factory=list)
    post_condition: Optional[LoopCondition] = None


@dataclass
class ExitStatement(ASTNode):
    exit_type: ExitType


@dataclass
class OnErrorResumeNextStatement(ASTNode):
    """On Error Resume Next - continue execution after errors."""

    pass


@dataclass
class OnErrorGoToStatement(ASTNode):
    """On Error GoTo 0 - reset error handling to default."""

    label: int  # 0 for resetting to default, or line number for GoTo


@dataclass
class Program(ASTNode):
    statements: List[ASTNode] = field(default_factory=list)


# Type alias for expressions
Expression = Union[
    NumberLiteral,
    StringLiteral,
    BooleanLiteral,
    NothingLiteral,
    EmptyLiteral,
    NullLiteral,
    Identifier,
    BinaryExpression,
    UnaryExpression,
    ComparisonExpression,
    MemberAccess,
    FunctionCall,
    MethodCall,
    NewExpression,
    ArrayAccess,
]

# Type alias for statements
Statement = Union[
    DimStatement,
    AssignmentStatement,
    SetStatement,
    PropertyAssignmentStatement,
    CallStatement,
    ExpressionStatement,
    IfStatement,
    SelectCaseStatement,
    ForStatement,
    ForEachStatement,
    WhileStatement,
    DoLoopStatement,
    ExitStatement,
    SubStatement,
    FunctionStatement,
    OnErrorResumeNextStatement,
    OnErrorGoToStatement,
    ReDimStatement,
    EraseStatement,
]
