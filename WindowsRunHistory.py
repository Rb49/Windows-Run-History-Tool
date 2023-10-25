import winreg


def sort_history_reg() -> tuple or None:
    """
    returns the history list and sorting key used on ot
    :return: tuple: [0]: list of history values. [1]: value of MRUList sub-key. None if not successful
    """
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    hive = winreg.HKEY_CURRENT_USER

    try:
        key = winreg.OpenKey(hive, key_path)
        info = {}
        sort_key = None

        #
        for i in range(winreg.QueryInfoKey(key)[1]):
            name, value, _ = winreg.EnumValue(key, i)

            if len(name) == 1 and name.isalpha():
                info[name] = value[:-2]
            elif name == "MRUList":
                sort_key = value

        # sort the list by order
        sorted_info = []
        for i in sort_key:
            sorted_info.append(info[i])

        winreg.CloseKey(key)
        return sorted_info, sort_key

    except FileNotFoundError:
        print(f"Registry key '{key_path}' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def add_item(item: str) -> True or None:
    """
    adds an item to the registry and updates the key
    :param item: str: value to be added to the registry
    :return: True is successful, None if not
    """
    global history, key

    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    hive = winreg.HKEY_CURRENT_USER

    try:
        reg_key = winreg.OpenKey(hive, key_path, 0, winreg.KEY_WRITE)
        # get the lowest unused letter name
        value_name = ""
        for i in range(ord('a'), ord('z') + 1):
            if chr(i) not in key:
                value_name = str(chr(i))
                break

        # save registry
        winreg.SetValueEx(reg_key, value_name, 0, winreg.REG_SZ, str(item + r"\1"))

        # update key
        key = str(value_name + key)
        winreg.SetValueEx(reg_key, "MRUList", 0, winreg.REG_SZ, key)
        winreg.CloseKey(reg_key)

        return True

    except FileNotFoundError:
        print(f"Registry key '{key_path}' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def delete_item(item: str) -> True or None:
    """
    deletes an item from the registry and updates the key
    :param item: str: value to be deleted from the registry
    :return: True is successful, None if not
    """
    global history, key

    key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\RunMRU"
    hive = winreg.HKEY_CURRENT_USER

    try:
        reg_key = winreg.OpenKeyEx(hive, key_path, 0, winreg.KEY_WRITE)

        # delete registry
        index = history.index(item)
        winreg.DeleteValue(reg_key, key[index])

        # update key
        key = str(key[:index] + key[index + 1:])
        winreg.SetValueEx(reg_key, "MRUList", 0, winreg.REG_SZ, key)
        winreg.CloseKey(reg_key)

        return True

    except FileNotFoundError:
        print(f"Registry key '{key_path}' not found.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def main() -> None:
    """
    main function running the ui loop indefinitely until broken
    :return: None
    """
    global history, key
    history, key = sort_history_reg()
    if not history:
        return

    while True:
        print(
            "\nWelcome to Windows Run History Tool!\nChoose an option:\n\t1. Show history\n\t2. Add item to history\n\t3. Delete item from history\n\t4. Quit\n")
        ans = input("> ")

        if ans.isdigit():
            # show history
            if int(ans) == 1:
                print("\nHistory items:")
                for i, value in enumerate(history):
                    print(f"\t{i + 1}. {value}")

            # add item
            elif int(ans) == 2:
                if len(history) >= 26:
                    print("Cannot add any more items to the history list.")
                    continue

                item_to_add = input("Write the item you want to add to the list.\n> ")
                if item_to_add in history:
                    print(f"Cannot add {item_to_add} to the history list.")
                    continue

                if add_item(item_to_add):
                    print("Addition was successful!")
                else:
                    print("Addition was not successful.")
                # update history
                history, key = sort_history_reg()

            # delete item
            elif int(ans) == 3:
                print("History items:")
                for i, value in enumerate(history):
                    print(f"\t{i + 1}. {value}")

                item_to_delete = input("Write the item you want to delete from the list.\n> ")
                if item_to_delete not in history:
                    print(f"Cannot delete '{item_to_delete}' from the history list.")
                    continue

                if delete_item(item_to_delete):
                    print("Deletion was successful!")
                else:
                    print("Deletion was not successful.")
                # update history
                history, key = sort_history_reg()

            # quit
            elif int(ans) == 4:
                return

        print("\nInvalid option. Please try again!\n")


# code starts here
if __name__ == "__main__":
    try:
        history, key = None, None
        main()
    except Exception as e:
        print(str(e))
    finally:
        quit()
