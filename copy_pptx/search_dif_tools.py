import copy
import subprocess
from typing import Union


def filter_lines(input_string):
    lines = input_string.split('\n')
    filtered_lines = [line for line in lines if line.startswith('+') or line.startswith('-')]

    plus_lines = []
    minus_lines = []

    for filtered_line in filtered_lines:
        if filtered_line.startswith('+'):
            plus_lines.append(filtered_line)
        else:
            minus_lines.append(filtered_line)

    final_plus_lines = []
    final_minus_lines = copy.deepcopy(minus_lines)
    for plus_line in plus_lines:
        not_plus_line = plus_line[1:]
        correct = True
        for minus_line in minus_lines:
            not_minus_line = minus_line[1:]
            if not_plus_line == not_minus_line:
                correct = False
                break
        if correct:
            final_plus_lines.append(plus_line)
        else:
            final_minus_lines.remove("-" + not_plus_line)

    output_string = '\n'.join(final_plus_lines + final_minus_lines)
    return output_string


def dif_dir(path_1, path_2):
    run = subprocess.run(["git", "diff", "--no-index", "--", path_1, path_2], capture_output=True)
    string = filter_lines(run.stdout.decode()).strip()
    if string:
        print(string + "\n")
