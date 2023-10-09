def remove_after_open_parenthesis(input_string):
    if input_string is not None:
        position = input_string.find(' (')
        return input_string[:position] if position != -1 else input_string
    else:
        return None


def convert_to_inches(feet_inches_str):
    if feet_inches_str is not None:
        # Remove any unwanted characters
        cleaned_str = feet_inches_str.replace("\\", "").replace("\"", "")

        # Split based on the single quote to separate feet and inches
        split_str = cleaned_str.split("'")

        # Initialize feet and inches to 0
        feet = 0
        inches = 0

        # Determine if the measurement includes feet, inches, or both
        if len(split_str) == 2:
            # Both feet and inches may be present
            feet = int(split_str[0])
            # Check if inches part is empty after stripping
            if split_str[1].strip():
                inches = int(split_str[1].strip())
        elif "\"" in feet_inches_str:
            # Only inches are present
            inches = int(cleaned_str)
        elif "'" in cleaned_str:
            # Only feet are present
            feet = int(split_str[0])
        else:
            # Invalid format, you may raise an exception or return None
            return None

        # Calculate the total inches
        total_inches = (feet * 12) + inches

        return str(total_inches)
    else:
        return None


def check_duplicate_dict(target_dict, dict_list):
    return target_dict in dict_list
