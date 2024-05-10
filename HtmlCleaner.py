# ラジオボタン向けhtml整形
import re
import datetime

def sort_att(line: str):
    # type > style > class > id > name > else
    pattern = r"<input\s+[^>]*?>"
    context = re.search(pattern, line)
    if context:
        # print(context.group(0))
        str_list = (context.group(0)[7:-1]+' ').split('" ')
        print(str_list)
    
        # listをdictに転換
        att_dict = {}
        for att in str_list:
            if att.strip():
                k, v = att.split('=')
                att_dict[k.strip()] = v.strip()
        
        res_str = "<input"
        for k in ["type", "style", "class", "id", "name"]:
            if k in att_dict:
                res_str += f" {k}={att_dict[k]}\""
                del att_dict[k]

        # 残り属性
        for k, v in att_dict.items():
            res_str += f" {k}={att_dict[k]}\""
        
        unmatched_parts = re.split(pattern, line)

        return f"{unmatched_parts[0]}{res_str}>{unmatched_parts[1]}"

    else:
        return line

def copy_file(input_file, output_file):
    try:
        with open(input_file, 'r', encoding="utf8") as f_input:
            with open(output_file, 'a+', encoding="utf8") as f_output:
                for line in f_input:
                    f_output.write(sort_att(line))
        print("ファイルの転写が完了しました。")
    except FileNotFoundError:
        print("指定されたファイルが見つかりません。")

def copy_file_id(input_file, output_file):
    try:
        name = ""
        resent_id = ""
        resent_index = 1
        name_list = []
        with open(input_file, 'r', encoding="utf8") as f_input:
            with open(output_file, 'a+', encoding="utf8") as f_output:
                for line in f_input:
                    # tdを目印にする
                    if "<td" in line:
                        name = ""
                        resent_index = 1
                    elif 'type="radio"' in line:
                        # name不一致の場合warning
                        _name = re.search(r'name="[^"]*?"', line).group(0)[6:-1]
                        if name:
                            if name != _name:
                                print(f"Warning! name={name} but also name={_name}")
                        else:
                            name = _name
                            if name in name_list:
                                print(f"Warning! name={name} duplicat")
                            else:
                                name_list.append(name)
                        
                        # idを配る
                        _id = re.search(r'id="[^"]*?"', line).group(0)[4:-1]
                        # print(f'id old:{_id}')
                        resent_id = f'r-{name}-{resent_index}'
                        line = line.replace(f'id="{_id}"', f'id="{resent_id}"')
                        resent_index += 1
                        # print(f'resent_id = {resent_id}')

                    elif '<label' in line:
                        if resent_id:
                            # print(f'bef:{line}')
                            _for = re.search(r'for="[^"]*?"', line).group(0)[5:-1]
                            line = line.replace(f'for="{_for}"', f'for="{resent_id}"')
                            # idを解放する
                            resent_id = ""
                            # print(f'aft:{line}')
                            # print(f'for = {_for}')

                    f_output.write(line)
        print("ファイルの転写が完了しました。")
    except FileNotFoundError:
        print("指定されたファイルが見つかりません。")

if __name__ == "__main__":
    # line = '									<input type="text" maxlength="3" value="111" class="postal-code-1"><span class="span-blank">-</span>'
    # print(line)
    # r = sort_att(line)
    # print(r)

    filepath = r"C:\ning\dev\mockup"
    filename = "MDA0110_届出事項登録_管理組合.html"
    print(f"処理開始：{filepath}\\{filename}")

    out_filename = filename[:-5]+datetime.datetime.now().strftime('%H%M%S')+filename[-5:]
    if False:
        # inputの属性を並び替え
        copy_file(filepath+'\\'+filename, filepath+'\\'+out_filename)
    else:
        # radioのid配り
        copy_file_id(filepath+'\\'+filename, filepath+'\\'+out_filename)

    print(f"処理完了：{filepath}\\{out_filename}")
