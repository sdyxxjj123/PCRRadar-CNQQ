import pandas as pd
import time
from datetime import datetime, timedelta
import pytz
from openpyxl.styles import Font
import os

def has_claimed_reward(login_timestamp):
    # 设置东八区时区
    tz = pytz.timezone('Asia/Shanghai')
    login_time = datetime.fromtimestamp(login_timestamp, tz)

    # 当前时间（东八区）
    now = datetime.now(tz)

    # 计算今天 5 点刷新时间
    today_reset_time = tz.localize(datetime(now.year, now.month, now.day, 5, 0, 0))

    # 如果当前时间还没到今天 5 点，则基准点回退一天
    if now < today_reset_time:
        today_reset_time -= timedelta(days=1)

    yesterday_reset_time = today_reset_time - timedelta(days=1)

    # 判断登录状态并返回数值
    if login_time >= today_reset_time:
        return 0  # 今天已登录
    elif login_time >= yesterday_reset_time:
        return 1  # 昨天已登录，但今天未登录
    else:
        return 2  # 昨天和今天都未登录
    
def get_KRANK(EXP):
    cur_dir = os.path.dirname(__file__)
    excel_path = os.path.join(cur_dir, 'CSV.xlsx')
    df = pd.read_excel(excel_path, header=0)
    
    levels = df.iloc[:, 0].values
    required_exps = df.iloc[:, 1].values

    current_level = levels[0]

    for i in range(len(required_exps)):
        if EXP >= required_exps[i]:
            current_level = levels[i]
        else:
            break

    # 判断是否已满级
    if current_level == levels[-1]:
        return current_level, None
    else:
        next_level_exp = required_exps[levels.tolist().index(current_level) + 1]
        return current_level, next_level_exp - EXP

def SY_data(res4,res5,sheet,uid,data,numx):
                    Krank,KrankEXP = get_KRANK(int(res4["princess_knight_rank_total_exp"]))

                    timeStamp = res4["last_login_time"]
                    timeArray = time.localtime(timeStamp)
                    otherStyleTime = time.strftime("%Y--%m--%d %H:%M:%S", timeArray)
                    if ((res5['quest_info']['talent_quest'][0]['clear_count'])%10) == 0 and (res5['quest_info']['talent_quest'][0]['clear_count']) !=0:
                        SYH = f"{((res5['quest_info']['talent_quest'][0]['clear_count'])//10)}-10"
                    else:
                        SYH = f"{((res5['quest_info']['talent_quest'][0]['clear_count'])//10)+1}-{(res5['quest_info']['talent_quest'][0]['clear_count'])%10}"
                    if ((res5['quest_info']['talent_quest'][1]['clear_count'])%10) == 0 and (res5['quest_info']['talent_quest'][1]['clear_count']) !=0:
                        SYS = f"{((res5['quest_info']['talent_quest'][1]['clear_count'])//10)}-10"
                    else:
                        SYS = f"{((res5['quest_info']['talent_quest'][1]['clear_count'])//10)+1}-{(res5['quest_info']['talent_quest'][1]['clear_count'])%10}"
                    if ((res5['quest_info']['talent_quest'][2]['clear_count'])%10) == 0 and (res5['quest_info']['talent_quest'][2]['clear_count']) !=0:
                        SYF = f"{((res5['quest_info']['talent_quest'][2]['clear_count'])//10)}-10"
                    else:
                        SYF = f"{((res5['quest_info']['talent_quest'][2]['clear_count'])//10)+1}-{(res5['quest_info']['talent_quest'][2]['clear_count'])%10}"
                    if ((res5['quest_info']['talent_quest'][3]['clear_count'])%10) == 0 and (res5['quest_info']['talent_quest'][3]['clear_count']) !=0:
                        SYG = f"{((res5['quest_info']['talent_quest'][3]['clear_count'])//10)}-10"
                    else:
                        SYG = f"{((res5['quest_info']['talent_quest'][3]['clear_count'])//10)+1}-{(res5['quest_info']['talent_quest'][3]['clear_count'])%10}"
                    if ((res5['quest_info']['talent_quest'][4]['clear_count'])%10) == 0 and (res5['quest_info']['talent_quest'][4]['clear_count']) !=0:
                        SYA = f"{((res5['quest_info']['talent_quest'][4]['clear_count'])//10)}-10"
                    else:
                        SYA = f"{((res5['quest_info']['talent_quest'][4]['clear_count'])//10)+1}-{(res5['quest_info']['talent_quest'][4]['clear_count'])%10}"
                    sheet[f'A{numx}'] = uid
                    if data == '0':
                        sheet[f'B{numx}'] = '/'
                    else:
                        sheet[f'B{numx}'] = data["uid"]
                    sheet[f'C{numx}'] = res4["user_name"]
                    sheet[f'D{numx}'] = res4["team_level"]
                    sheet[f'E{numx}'] = res4["total_power"]
                    sheet[f'F{numx}'] = res4["unit_num"]
                    sheet[f'G{numx}'] = Krank
                    sheet[f'H{numx}'] = KrankEXP
                    sheet[f'I{numx}'] = otherStyleTime
                    sheet[f'J{numx}'] = SYH
                    sheet[f'K{numx}'] = SYS
                    sheet[f'L{numx}'] = SYF
                    sheet[f'M{numx}'] = SYG
                    sheet[f'N{numx}'] = SYA
                    if has_claimed_reward(timeStamp) == 0:
                        pass
                    elif has_claimed_reward(timeStamp) == 1:
                        cell = sheet[f'I{numx}']
                        cell.font = Font(color="800080")
                    else:
                        cell = sheet[f'I{numx}']
                        cell.font = Font(color="FF0000")