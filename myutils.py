import datetime
import os
import pickle

import myconstants


def save_param(param_name, param_value, filename="saved.pkl"):
    params_dict = get_all_params(filename="saved.pkl")
    params_dict[param_name] = param_value
    
    with open(filename, 'wb') as file_handle:
        pickle.dump(params_dict, file_handle)

def get_all_params(filename="saved.pkl"):
    if not os.path.isfile(filename):
        with open(filename, 'wb') as file_handle:
            pickle.dump({"<ДатаВремя создания>": datetime.datetime.now()}, file_handle)

    with open(filename, "rb") as file_handle:
        params_dict = pickle.load(file_handle)

    return(params_dict)

def load_param(param_name, default="", filename="saved.pkl"):
    params_dict = get_all_params(filename="saved.pkl")
    
    return(params_dict.get(param_name, default))

def get_files_list(path2files="", files_starts="", files_ends=".xlsx"):
    files_list = \
        [one_file[len(files_starts):][:-len(files_ends)] \
            for one_file \
                in os.listdir(path2files) \
                    if (one_file.lower().startswith(files_starts.lower()) and \
                        one_file.lower().endswith(files_ends.lower()))]
    
#    return(sorted(files_list, reverse=True))
    return(files_list)


if __name__ == "__main__":
#    param_name = myconstants.LAST_REPORT_PARAM_NAME
#    print(load_param(param_name, "<пусто>"))
#    save_param(param_name, 7)
#    print(load_param(param_name, "<пусто>"))
    print(get_files_list("RawData", "", ".xlsx"))
