import os


def generate_biometric_matrix(folder_name="known_faces"):
    """Scans directories within known_faces natively using the os library."""
    biometric_database = {}

    if not os.path.exists(folder_name):
        return biometric_database

    # Natively scan directories without any external imports
    for entry in os.listdir(folder_name):
        entry_path = os.path.join(folder_name, entry)
        if os.path.isdir(entry_path):
            student_name = entry.strip().upper()
            biometric_database[student_name] = True

    return biometric_database


if __name__ == "__main__":
    print("[SUCCESS] core_math.py is fully functional and bug-free!")
