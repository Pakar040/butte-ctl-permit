def remove_after_open_parenthesis(input_string):
    position = input_string.find(' (')
    return input_string[:position] if position != -1 else input_string


def convert_to_inches(feet_inches_str):
    if feet_inches_str is not None:
        feet, inches = map(int, feet_inches_str.replace("\\", "").replace("\"", "").split("' "))
        total_inches = (feet * 12) + inches
        return str(total_inches)
    else:
        return None


def check_duplicate_dict(target_dict, dict_list):
    return target_dict in dict_list
