import ast
import json

def get_function_lines(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        source = f.read()
    
    tree = ast.parse(source)
    functions = {}
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) or isinstance(node, ast.AsyncFunctionDef):
            functions[node.name] = [node.lineno, node.end_lineno]
        elif isinstance(node, ast.ClassDef):
            functions[node.name] = [node.lineno, node.end_lineno]
    return functions

if __name__ == "__main__":
    funcs = get_function_lines('app.py')
    print(json.dumps(funcs, indent=2))
