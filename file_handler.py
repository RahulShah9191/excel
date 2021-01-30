import string
import platform
import tempfile
import shutil
import subprocess
import re
import paramiko




def is_file_present(abs_path_of_file):
    if os.path.exists(abs_path_of_file) and os.path.isfile(abs_path_of_file):
        return True
    else:
        return False


def is_dir_present(abs_path_of_dir):
    if os.path.exists(abs_path_of_dir) and os.path.isdir(abs_path_of_dir):
        return True
    else:
        return False


def get_file_data(abs_path_of_file, remove_binary_chars=False):
    content = None
    data = ""
    try:
        with open(file=abs_path_of_file, mode='r', encoding="utf-8", errors='ignore') as f:
            if not remove_binary_chars:
                content = f.read()
            else:
                content = ""
                for c in f.read():
                    if c in string.printable:
                        content += c
                        continue
    except Exception as e:
        data = ""
        print(abs_path_of_file + " | Unable to read/open file! \n Exception :: {e}".format(e=e))
    finally:
        data = content.replace('\r\n', '\n')
    return data


def write_data_to_file(data, abs_path_of_file, mode, overwrite_if_exist=False):
    status = False
    if not is_file_present(abs_path_of_file) or (is_file_present(abs_path_of_file) and overwrite_if_exist):
        try:
            with open(file=abs_path_of_file, mode=mode, encoding="utf-8", errors='ignore') as f:
                f.write(data)
                status = True
        except Exception as e:
            status = False
            print("Unable to read/open file - {f}! \n Exception :: {e}".format(f=abs_path_of_file, e=e))
    else:
        status = False
        LOGGER.error("File {f} already exists and overwrite is set to False!".format(f=abs_path_of_file))
    return status


def change_file_permissions(path_to_file, permissions="777"):
    if is_file_present(path_to_file):
        try:
            if platform.system() in ("Linux", "Darwin"):
                os.chmod(path_to_file, int("0o"+permissions))
                return True
            elif platform.system() == "Windows":
                return True
        except Exception as e:
            print("Unable to change permissions!! \n Exception :: "+e)
    else:
        return False


def delete_file_if_exists(path_to_file):
    if is_file_present(path_to_file):
        try:
            os.remove(path_to_file)
            print("File deleted :: {0}".format(path_to_file))
            return True
        except Exception as e:
            print("Unable to delete file! \n Exception :: {0}".format(e))
    else:
        print("File does not exists! :: {0}".format(path_to_file))
        return None


def get_tempdir():
    return tempfile.gettempdir()


def create_temp_file_with_data(file_name, data=None, dir_name=tempfile.gettempdir()):
    if is_dir_present(dir_name) and str(dir_name).strip(" ") != "" and not is_file_present(file_name):
        try:
            f = open(file_name, "a+")
            if data is not None:
                f.write(data)
            f.close()
            status = True
        except Exception as e:
            print("Unable to create temporary file {f}".format(f=file_name))
            status = False
    else:
        LOGGER.error("Unable to create file {f} in directory {d}!".format(f=file_name, d=dir_name))
        status = False
    return status

def copy_file_from_server(src_server_ip,src_server_user,src_server_pwd,src_abs_path_to_file,tmp_file):
    _src_server_connection = ssh_lib.SSH(str(src_server_ip), str(src_server_user), str(src_server_pwd))
    _src_server_connection.copy_from_server_to_local(src_abs_path_to_file, tmp_file)
    _src_server_connection.close_connection()


def get_file_from_server_or_local(src_abs_path_to_file, src_server_ip, src_server_user, src_server_pwd,
                                  tgt_abs_path_to_file=None, tgt_server_ip=None, tgt_server_user=None, tgt_server_pwd=None):
    """
    :param src_abs_path_to_file: Requires absolute path to src file ( on local or server )
    :param src_server_ip:
    :param src_server_user:
    :param src_server_pwd:
    :param tgt_abs_path_to_file:
    :param tgt_server_ip:
    :param tgt_server_user:
    :param tgt_server_pwd:
    :return:
    """
    status = False
    tmp_file = os.path.abspath(os.path.join(get_tempdir(), string_lib.random_string_generator(10) + ".tmp"))
    print("tmp_file :: {tf}".format(tf=tmp_file))
    try:
        local_ip = os_lib.get_ip_address()
        src_dir = os.path.dirname(src_abs_path_to_file)
        LOGGER.debug(f"local ip = {local_ip} and src_server_ip = {src_server_ip}")
        LOGGER.debug(f"src_dir = {src_dir} and tmp_file = {tmp_file}")

        if local_ip == src_server_ip and str(src_server_ip).strip(" ") != "" and local_ip != "".strip(" "):
            # in case python code is running on the same server!
            if docker_lib.check_inside_docker():
                LOGGER.debug("Docker running on same machine.\n Process inside docker.")
                copy_file_from_server(src_server_ip,src_server_user,src_server_pwd,src_abs_path_to_file,tmp_file)
                status = True
            else:
                if is_dir_present(src_dir) and src_dir is not None and str(src_dir).strip(" ") != "":
                    # checking for absolute path for log_file and verifying directory
                    if is_file_present(src_abs_path_to_file):
                        # checking if file is present, if yes copying to local storage for using
                        if str(src_abs_path_to_file).strip(" ") != str(tgt_abs_path_to_file).strip(" "):
                            shutil.copy(src_abs_path_to_file, tmp_file)
                            # shutil.copystat(src_abs_path_to_file, tmp_file)
                            print("File copied from server to local.")
                            status = True
                        else:
                            # checking if src and tgt files have the same name, to be copied in same server!
                            LOGGER.error("Source and target files are on the same server. They can not have same names!")
                            status = False
                    else:
                        LOGGER.error("{t} | Unable to locate AdServer log file {f}".format(t=TAG, f=src_abs_path_to_file))
                        status = False
                else:
                    LOGGER.error("{t} | Unable to locate AdServer log directory {d}".format(
                        t=TAG, d=os.path.dirname(src_abs_path_to_file)))
        else:
            # in case if python code is running in a different server!
            copy_file_from_server(src_server_ip, src_server_user, src_server_pwd, src_abs_path_to_file, tmp_file)
            status = True
    except Exception as e:
        print("{tag} | Unable to access file {f}".format(tag=TAG, f=src_abs_path_to_file))
        status = False

    if tgt_abs_path_to_file is None or str(tgt_abs_path_to_file).strip(" ") == "":
        tgt_abs_path_to_file = tmp_file

    try:
        # check if tgt server == local server, if yes then just copy file from tmp location to target location
        if tgt_server_ip == local_ip or tgt_server_ip is None or str(tgt_server_ip).strip(" ") == "" and status:
            try:
                # shutil.copystat(tmp_file, tgt_abs_path_to_file)
                shutil.move(tmp_file, tgt_abs_path_to_file)
                status = True
            except Exception as e:
                print("Unable to move tmp file {tmp} to {tgt}".format(tmp=tmp_file, tgt=tgt_abs_path_to_file))
                status = False
        # else if, tgt server != local and tgt server, username and password are given, then copy file to server
        elif tgt_server_ip != local_ip and tgt_server_ip is not None and tgt_server_user is not None \
                and tgt_server_pwd is not None and status:
                _tgt_server_connection = ssh_lib.SSH(str(tgt_server_ip), str(tgt_server_user), str(tgt_server_pwd))
                _tgt_server_connection.copy_from_local_to_server(tmp_file, tgt_abs_path_to_file)
                _tgt_server_connection.close_connection()
                print("File copied from local to server.")
                status = True
        else:
            LOGGER.error("Unable to copy file {tmp_file} to {server_file}".format(
                tmp_file=tmp_file, server_file=tgt_abs_path_to_file))
            status = False
    except Exception as e:
        print("Unable to copy file {tmp_file} to {server_file}".format(
            tmp_file=tmp_file, server_file=tgt_abs_path_to_file))
        status = False

    if status:
        return tgt_abs_path_to_file
    else:
        return status


def update_remote_file(searched_string,replaced_string,filepath,hostname=app_config.TRANS_IP,username=app_config.TRANS_USERNAME,password=app_config.TRANS_PASSWORD):
    ssh_this=ssh_lib.SSH(hostname,username,password)
    sftp= ssh_this.ssh.open_sftp()
    file_open = sftp.open(filepath, mode='rb')
    try:
        data = file_open.read()
        print(f"type of data is {type(data)} \n type of search string {searched_string} \n type of replace string is {type(replaced_string)}")
        file_open.close()
        data = re.sub(searched_string.encode(),
                      replaced_string.encode(), data)

        file_open = sftp.open(filepath, mode='w')
        file_open.write(data)
        file_open.close()
        ssh_this.close_connection()
    finally:
        file_open.close()
        ssh_this.close_connection()


def pre_process_tcp_dump(tcp_dump_file, pattern, groups):
    data = get_file_data(tcp_dump_file)
    df = regex_lib.get_regex_groups_in_df(regex=pattern, data=data, group_names=groups)
    return df



