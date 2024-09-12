import antlr4

from VisualBasic6Lexer import VisualBasic6Lexer
from VisualBasic6Parser import VisualBasic6Parser

VB6_CODE = """
Function ThisIsObjectFunction() As Object
End Function

Function FindSqrt(radicand As Single) As Single
End Function

Private Sub Command1_Click()
    Dim a As String
    a = ThisIsObjectFunction()

    Dim b, c As Single
    b = FindSqrt(c)

End Sub
    """


def find_node(root, cls):
    if hasattr(root, "children"):
        for child in root.children:
            if isinstance(child, cls):
                return child
            c = find_node(child, cls)
            if c is not None:
                return c


def find_nodes(root, cls):
    if hasattr(root, "children"):
        for child in root.children:
            if isinstance(child, cls):
                yield child
            yield from find_nodes(child, cls)


def get_name(node):
    return find_node(node, VisualBasic6Parser.AmbiguousIdentifierContext)


def get_all_names(node):
    return list(find_nodes(node, VisualBasic6Parser.AmbiguousIdentifierContext))


def get_type(node):
    as_type_node = find_node(node, VisualBasic6Parser.AsTypeClauseContext)
    type_nodes = find_nodes(as_type_node, VisualBasic6Parser.BaseTypeContext)
    return next(type_nodes, None)


def get_all_types(node):
    return list(find_nodes(node, VisualBasic6Parser.BaseTypeContext))


def get_all_functions(node):
    return list(find_nodes(node, VisualBasic6Parser.FunctionStmtContext))


def get_all_vars(node):
    return list(find_nodes(node, VisualBasic6Parser.VariableStmtContext))


def get_all_used_var_locations(var_node):
    var_names = [
        get_name(node).getText()
        for node in find_nodes(var_node, VisualBasic6Parser.VariableSubStmtContext)
    ]
    for used_var_let_stmt in find_nodes(
        var_node.parentCtx.parentCtx, VisualBasic6Parser.LetStmtContext
    ):
        ident = get_name(used_var_let_stmt)
        value_stmt = find_node(used_var_let_stmt, VisualBasic6Parser.ValueStmtContext)
        if ident.getText() in var_names:
            yield used_var_let_stmt, value_stmt


def report_function_locations(all_functions):
    print("Functions:")
    for func in all_functions:
        print(f"{get_name(func).getText()}: {get_type(func).getText()}")
    print()


def report_var_locations(all_vars):
    print("Variables:")
    for v in all_vars:
        print(v.getText())
    print()


def report_object_assignment(all_vars, all_functions):
    function_dict = {}
    for func in all_functions:
        function_dict[get_name(func).getText()] = get_type(func).getText()

    print("Result:")

    for v in all_vars:
        used_locations = get_all_used_var_locations(v)
        for lhs, rhs in used_locations:
            for name in get_all_names(rhs):
                if (
                    name.getText() in function_dict
                    and function_dict[name.getText()] == "Object"
                ):
                    print(lhs.getText())


def main():
    input_stream = antlr4.InputStream(VB6_CODE)
    lexer = VisualBasic6Lexer(input_stream)
    stream = antlr4.CommonTokenStream(lexer)
    parser = VisualBasic6Parser(stream)
    tree = parser.startRule()

    all_functions = get_all_functions(tree)
    all_vars = get_all_vars(tree)

    report_function_locations(all_functions)
    report_var_locations(all_vars)
    report_object_assignment(all_vars, all_functions)


if __name__ == "__main__":
    main()
