import string
import random
import uuid
from urllib.parse import urlparse, urlunparse, parse_qs
from requests.models import PreparedRequest



def random_string_generator(size=6, chars=string.ascii_uppercase + string.digits):
    """ Return the randon string generated
        Params:
        -------
        size : int, default 6
            The size if the string
        chars : Char set, default ASCII Uper case
            The character formate whihc need to generate
    """
    return ''.join(random.choice(chars) for _ in range(size))


def generate_uuid():
    """Return  the Random UUID generated
    """
    return uuid.uuid1()


def remove_space(input_str):
    """ Remove the extra spaces from string
        Params:
        input_str : str
            Then string in which extra spaces needs to remove
    """
    return " ".join(input_str.strip().split())


def is_string_empty_or_none(*args):
    result=[]
    for arg in args:
        if arg is None or str(arg).strip(" ") == "":
            result.append(True)
        else:
            result.append(False)
    result = False if False in result else True
    return result


def get_params_from_url(url):
    try:
        if not is_string_empty_or_none(url):
            u = urlparse(url)
            qs = parse_qs(u.query)
            return qs
        else:
            print(f"Url can not be empty or None! Check url - {url}")
    except Exception as e:
        print(f"Unable to parse query string from url - {url}. \nException:: {e}")
        return False


def add_params_to_url(url, params):
    out_url = ""
    try:
        req = PreparedRequest()
        req.prepare_url(url, params)
        out_url = req.url
    except Exception as e:
        print(f"\n Exception occured: {e}!! \n URL: {url} \n PARAMS: {params}")
    return out_url


def drop_params_from_url(url, params, drop_keys, assert_param_is_present=False):
    out_url = ""
    try:
        pre_req = PreparedRequest()
        pre_req.prepare_url(url, params)
        u = urlparse(pre_req.url)
        current_params = parse_qs(u.query)

        if assert_param_is_present:
            keys_diff = set(drop_keys) - set(current_params)
            assert keys_diff == set(), f"Keys not present : {keys_diff}"

        final_params = {k: v for k, v in current_params.items() if k not in drop_keys}
        current_url = u._replace(params="")._replace(query="")._replace(fragment="").geturl()
        pre_req.prepare_url(current_url, final_params)
        out_url = pre_req.url
    except Exception as e:
        print(f"\n Exception occured: {e}!! \n URL: {url} \n PARAMS: {params}")
    return out_url


def get_string_from_dict(input_dict, keys_delimiter, key_value_sep):
    # print("input_string==", input_string)
    try:
        output_str = f"{keys_delimiter}".join(list(map(lambda x: f"{x[0]}{key_value_sep}{x[1]}", input_dict.items())))
    except:
        output_str = {}
    return output_str


def get_dict_from_string(input_string, keys_delimiter, key_value_sep):
    # print("input_string==", input_string)
    try:
        output_dict = dict(map(lambda x: x.split(key_value_sep), input_string.split(keys_delimiter)))
    except:
        output_dict = {}
    return output_dict


def get_url_for_tcp_dump_qs(url, qs, update_params):
    d = get_params_from_url('http://xyz.com?'+qs)
    d.update(update_params)
    u = add_params_to_url(url=url, params=d)
    return u


def get_url_list(url, qs_list, update_params):
    out_list = []
    for qs in qs_list:
        out_list.append(get_url_for_tcp_dump_qs(url=url, qs=qs, update_params=update_params))
    return out_list


def is_list_equal(list_1, list_2):
    if len(list_1) != len(list_2):
        return False
    else:
        for i,j in zip(sorted(list_1), sorted(list_2)):
            if i != j:
                return False
        return True


if __name__ == "__main__":
    # print(is_list_equal([],[]))
    pass
