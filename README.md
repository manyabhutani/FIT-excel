# FIT-excel
The task is to develop classes that simulate the backend of a spreadsheet processor. 
The interface and mandatory classes:

CSpreadsheet - the spreadsheet processor itself. An instance represents a single sheet. The exact interface is described below. This class will be implemented by you.
CPos - an identifier of a cell in the sheet. The cells are identified in the standard way: a non-empty sequence of letters A to Z (that denotes the column) is followed by a non-negative number (that denotes a row, 0 is a valid row). Both upper and lower case letters are permitted (case insensitive). Columns are named A, B, C, ..., Z, AA, AB, AC, ..., AZ, BA, BB, ..., ZZZ, AAAA, ... Example cell identifiers are: A7, PROG7250, ... Consult real spreadsheet processor if unsure. The identifier in the textual form is human-friendly, however, it is not well suited for the implementation. Therefore, there exists class CPos that represents a cell identifier. The class is expected to parse the input string and store column/row values that are easier to work with. The implementation of CPos is your task.
CExpressionBuilder - an abstract class that is used by the expression syntax analyzer. If you decide to use the delivered syntax analyzer, you will have to implement a subclass of CExpressionBuilder.
parseExpression() - a function in the testing environment (and in the attached library) that does the syntax analysis and passes the parsed results to the your subclass of CExpressionBuilder.
CValue - represents a value of a cell. The value is one of: undefined, decimal number, or string. Class CValue is a named specialization of generic class std::variant.
further - your implementation will probably need many other classes and functions. These are not constrained by the testing environment.
Expressions in the cells:

A cell in the spreadsheet may contain various values: a cell may be empty (undefined), it may contain a decimal number, a string, or an expression. The expressions are alike the expressions in standard spreadsheet processors, however, the expressions are limited in the number of supported functions. An expression may contain:

a numeric literal - an integer or a decimal number with optional exponent, e.g., 15, 2.54, 1e+8, 1.23e-10, ...
a string literal - a string is a sequence of arbitrary characters in double quotes. If there is a double quote contained in the string, it must be doubled. There is no restriction on the length of the string, moreover, a newline character may be included in the string.
a reference to a cell - the reference uses standard notation (column: a sequence of letters, row - an integer). Both upper and lower case letters may be used (case insensitive). Examples are: A7, ZXCV789456, ... A reference may be either absolute, or relative. Moreover, absolute/relative reference may be set separately for columns and rows. An absolute reference refers to a cell, the reference is not modified when the cell is copied. On the other hand, a relative reference is updated when copying cells. An absolute reference is denoted by a dollar sign ($). Examples are: A5, $A5, $A$5, A$5.
cell range - a rectangular array of cells, the rectangle is determined by the upper left and lower right corner, separated by a colon (:). An example: A5:X17. Again, the cell references may be either absolute, or relative. Examples: A$7:$X29, A7:$X$29, ...
functions - an expression may invoke a function and pass parameters to the function. The following functions are supported:
sum (range) - the function evaluates all cells in the given range and sums all values that are numbers (it ignores undefined and string values). If there is not any numeric value in the given range, then the result of sum is undefined value.
count (range) - the function evaluates all cells in the given range and counts all cells with numeric or string values (it does not count undefined values).
min (range) - similar to sum, the result is the smallest numeric value in the range, or undefined value (no numeric values in the range).
max (range) - similar to sum, the result is the greatest numeric value in the range, or undefined value (no numeric values in the range).
countval (value, range) - counts the number of cells in the range that evaluate to value. Do not consider any epsilon tolerance when comparing numerical values.
if (cond, ifTrue, ifFalse) - the function evaluates the condition cond. If the result is undefined value, or a string, then the result of if is undefined value. If the condition evaluates to nonzero number, then the result of if is the value of expression ifTrue. If the condition evaluates to number zero, then the result of if is the value of expression ifFalse.
parentheses - are used to change the priority of subexpressions.
binary operators ^, *, /, and - are used to raise to some power, multiplication, division, and subtraction, respectively. All these operators require two numerical operands. If either operand is not a number, the result is undefined value. Next, division by zero results in undefined value. Do not consider any epsilon tolerance when checking for zero value in the divident.
operator + is similar to the other binary operators. Moreover, it stands for concatenation and concatenates the left and right operands if used for number+string, string+number, or string+string (numbers are converted to string using std::to_string function).
unary minus - changes the sign. The operator may only be used for numbers, it results in undefined value if its operand is a string or an undefined value.
relational operators < <=, >, >=, <>, = compare the operands. The operands must be either two numbers, or two strings. The result is number 1 (the relation holds) or 0 (the relation does not hold). If either operand is an undefined value, or if comparing number-to-string or string-to-number, then the result is an undefined value. Exact comparison is used for numbers (no epsilon tolerance).
The priority and associativity of the operators, from the highest priority:
function call,
power ^, left associativity,
unary -, right associativity,
multiplication * and division /, left associativity,
addition/concatenation + and subtraction -, left associativity,
relational operators < <=, > a >=, left associativity,
equality/inequality <> and =, left associativity.
The above explanation is important when implementing your own parser. The parser provided in function parseExpression() (in the attached library and in the testing environment) follows the above rules.

Class CPos

the class represents an identifier of a cell. The testing environment uses only the constructor, the constructor is called with the cell identifier in textual form, e.g., "A7". The implementation is expected to process and/or store the string. If the parameter is invalid, the constructor should throw std::invalid_argument exception. The internal representation and further interface of the class is left on you.

Function parseExpression() and class CExpressionBuilder

Function parseExpression( expr, builder ) represents the interface of the delivered expression parser. The parameter is the expression to parse and a builder object. The builder provides interface that is called by the parser as it analyzes the expression. The function either successfully parses the expression, or it fails and throws exception std::invalid_argument with the description of the problem. You must provide the builder object to use the function. The builder object is a subclass of abstract class CExpressionBuilder, the implementation of this subclass is left on you.

The parser function transforms the input expression from infix notation (e.g., 4 + A2 + B$7) into its postfix equivalent, i.e., 4 A2 + B$7 + in our example. The postfix form is not explicitly constructed, instead, the parser invokes the corresponding builder method whenever it identifies a value or a subexpression. The example would result in the following sequence of calls:

valNumber (4)
valReference ("A2")
opAdd ()
valReference ("B$7")
opAdd ()
The postfix notation is significantly easier to process. Use a stack data structure inside your builder subclass, push all delivered operands (valNumber, valString, valReference, ...) to this stack. Whenever there is an operator (opAdd, opSub, ..., funcCall), pop the corresponding number of operands from the stack and compute the new value (e.g., pop two operands and add them when opAdd is invoked). You have to be cautious, the left-hand operand is located deeper in the stack. Finally, push the computed result (e.g., the sum) to the stack. When the entire expression is successfully parsed, the stack contains exactly one value - the final result.

Indeed, the interface of the builder is very flexible. The above idea shows how to use the parser/builder to evaluate the expressions: store the resulting numbers/strings to the stack, the final result is the final value on the stack when the parsing finishes. Therefore, the value of a cell may be computed by repeatedly parsing the expression. While this technically works, the computation is rather slow. Cell X needs to be recomputed if its content changes, or if the value changes in any cell that is referenced from X. Parsing the string over and over is inefficient. You need to slightly modify the idea and use the parser/builder to construct AST tree. The (repeated) computation by means of an AST is much faster.

Abstract class CExpressionBuilder

opAdd()
apply addition/concatenation (+),
opSub()
apply subtraction (-),
opMul()
apply multiplication (*),
opDiv()
apply division (/),
opPow()
apply raise-to-power (^),
opNeg()
apply unary minus (-),
opEq()
apply comparison for equality (=),
opNe()
apply comparison for inequality (<>),
opLt()
apply comparison less-than (<),
opLe()
apply comparison less-than-or-equal (<=),
opGt()
apply comparison greater-than (>),
opGe()
apply comparison greater-than-or-equal (>=),
valNumber ( num )
add a new numerical operand, the parameter is the number to add,
valString ( str )
add a new string operand, the parameter is the string to add,
valReference ( str )
add a new reference-to-a-cell, the parameter is a string with the cell identifier (e.g., A7 or $x$12),
valRange ( str )
add a new range, the parameter is a string with the definition of the range (e.g., A7:$C9 or $x$12:z29),
funcCall ( fnName, paramCnt )
apply a function. The first parameter (fnName) is the name of the function, the second parameter (paramCnt) is the number of function parameters. The attached parser checks for valid function name (sum, min, max, count, if, and countval), moreover, it checks that the number of parameters and their types are compatible with the function (e.g., if requires 3 parameters and none of them may be a range, sum requires one parameter and the parameter must be a range, ...). Therefore, this validation is not needed in the builder implementation. Obviously the validation is needed if you decide to implement your own parser.
Class CSpreadsheet

default constructor
Creates an empty spreadsheet.
copy/move constructor/operator=, destructor
Support for the copy/move/destruct operations.
save ( os )
The method saves the contents of the spreadsheet into the given output stream. The saved data must be sufficient to restore the identical spreadsheet later, including the expressions. The method returns true (success) or false (failure).
load ( is )
The method reads the contents of a spreadsheet from an input stream. The result is either true (success) or false (failure, corrupted data).
setCell ( pos, contents )
The method sets content to cell pos. The contents is provided as a string, the following is permitted:
a valid number, e.g., setCell ( CPos ("A1"), "123456.789e-9" ),
a string with any characters, however, the string must not start with character =, e.g., setCell ( CPos ("A1"), "abc\ndef\\\"" ) (note that backslashes are needed to form a valid string in C++),
an expression that defines the value of the cell. The expression must start with character =, e.g., setCell ( CPos ("A1"), "=A5-($B7*c$4+sum($X9:AC$124))^2" ),
Return value of setCell is true (success), or false (a failure, an invalid expression, the cell remains unchanged). A call to setCell may change the values in other cells. However, we do not recommend to recompute all cells that depend on the modified value. Instead, recompute the values at later in getValue when the value is really needed.
getValue ( pos )
The method computes the value of the given cell, the computed value is the return value. If the value cannot be computed (the cell is not defined yet, there is a cyclic dependence in the computation), then the result is an undefined value CValue().
copyRect ( dstCell, srcCell, w, h )
The method copies the cells in the given rectangle to the destination. The source rectangle is defined by the upper left corner srcCell, the size of the rectangle is w columns and h rows, the destination upper left corner is dstCell. Caution: the source and destination rectangles may overlap.
capabilities ()
class method that reports a bit combination of supported optional features to test. The return value is any combination of constants SPREADSHEET_xxx, or 0, if you do not implement any optional feature. If you actually implement a feature and to not list it in capabilities(), then the feature is not tested and you will be awarder fewer points. On the other hand, if a feature is not implemented yet it is listed in capabilities, then the feature is tested and the test is going to fail. Moreover, if the test fails in a segfault or timeout, the further tests will be omitted (thus you may lose points if you correctly implement the features tested in the omitted tests). Therefore, it is in your best interest to list the actually implemented features to test. The constants are the following:
SPREADSHEET_CYCLIC_DEPS - your implementation handles cyclic dependencies among the cells,
SPREADSHEET_FUNCTIONS - your implementation supports functions in expressions,
SPREADSHEET_FILE_IO - your implementation handles file saving and loading even if the files are corrupted,
SPREADSHEET_SPEED - the computation of cell values is optimized, thus run the optional speed test,
SPREADSHEET_PARSER - you use your own parser, run the bonus test.
