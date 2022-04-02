# encoding: utf-8
import pandas as pd
import re
import itertools
import copy
import multiprocessing
import random


def array_headman(headmans):
    '''
    求出所有小组长的排列方式
    小组长所有排列方式个数为：A7,7
    :param headmans: 所有的小组长
    :return: 所有小组长的排列顺序
    '''
    headmans_number = headmans['人员']     # 所有小组长
    return list(itertools.permutations(headmans_number))


def read_file(path, sheet_name, col_names, group_settings):
    '''
    读取表格数据，并选出两个班长，七个小组长
    功能包括：
        1.读取excel表格中的数据
        2.清洗数据
        3.选出两名班长和七名组长
    班长与小组长选取规则为：
        1.班长：总分前两名
        2.组长：总分第三名到第九名
    :param path: 文件路径
    :param sheet_name: 需要读取的表格名
    :param col_names: 读取数据的所有列名
    :param group_settings: 分组设置
    :return:
    '''
    xls_datas = pd.read_excel(path, sheet_name=sheet_name, names=col_names)       # 读取excel表格
    col0_drop_null = xls_datas[~xls_datas['人员'].isnull()]
    # 清洗数据，去除多余数据
    col0_drop_null['0'] = col0_drop_null['人员'].str.contains(r'^本所人员', flags=re.IGNORECASE, regex=True)
    members = col0_drop_null[col0_drop_null['0'] == True]

    members['折算总分'] = members['折算总分'].astype(dtype='double')            # 将总分的数据类型改为'double'类型
    members = members.sort_values(by=['折算总分'], ascending=False)           # 根据总分对人员进行排序
    members.reset_index(drop=True, inplace=True)                 # 对清洗后的数据重新排序
    # 选择前两名为班长
    monitors = members.loc[[0,1]]
    members = members.drop([0,1])
    members.reset_index(drop=True, inplace=True)

    # 去除班长后，选择前七名为七个小组的小组长
    headmans = members.loc[range(group_settings['group_n'])]

    all_employees = members.drop(range(group_settings['group_n']))         # 所有待分组的员工
    headman_array = array_headman(headmans)                      # 求出七个小组长的所有排列情况
    return monitors, headman_array, all_employees, members            # 返回班长、小组长、所有待分组员工、所有人员数据


def group_other_members(all_group, other_members, headmans,
                        group_settings, all_members, all_fits_groups):
    '''
    将未分组的六个人员分配到每个小组
    所有的分配方式个数为：
        分配方式个数 = 迭代次数 * A(所有人员数-5*小组数),(所有人员数-5*小组数)
    :param all_group: 排除小组六个未分配人员的分组方式
    :param other_members: 为分组的六个人员
    :param headmans: 小组长所有的排列方式
    :param group_settings: 分组设置
    :param all_members: 所有成员的数据
    :param all_fits_groups: 存储所有满足条件的分组方式
    :return:
    '''
    all_groups = [[[] for item in range(7)]]
    all_groups = members_group1(other_members, all_groups)       # 加入其它六名未分配人员后的所有分组方式
    for groups_item in all_groups:
        for group_index, group_item in enumerate(groups_item):
            group_item.extend(all_group[group_index])          # 将其中一名未分组人员加入分组
        # 判断分组方式是否满足分组条件
        justice_groups(headmans, groups_item, group_settings,
                       all_members, all_fits_groups)


def group_deduplication(all_groups):
    '''
    对列表元素去重
    :param all_groups: 要去重的列表
    :return: 已去重的列表
    '''
    temp = list()
    for item in all_groups:
        if item not in temp:
            temp.append(item)
    return temp


def members_group(all_employees, iter_count, headmans,
                  group_settings, all_members, all_fits_groups):
    '''
    小组分配主函数
    功能包括：
        1.随机选择五个人作为一个小组；
        2.对未分组的人员进行分组；
    :param all_employees: 所有待分配的员工(去重班长和小组长)
    :param iter_count: 迭代次数
    :param headmans: 小组组长
    :param group_settings: 分组设置
    :param all_members: 所有人员数据
    :param all_fits_groups: 存储所有满足条件的分组
    :return:
    '''
    for i in range(iter_count):
        all_group = list()
        employees_indexs = [item_index for item_index,_ in enumerate(all_employees)]        # 所有员工索引(去除班长和小组长)
        for group_index in range(7):
            all_group_item = list()
            random_employee_index = random.sample(employees_indexs, 5)              # 随机抽取5名人员分成一个小组
            for random_index_item in random_employee_index:
                all_group_item.append(all_employees[random_index_item])            # 5名人员分组
                employees_indexs.remove(random_index_item)                     # 剔除该5名人员
            all_group.append(all_group_item)                                # 将7个小组结合，生成小组分配方式
        other_members = list()
        for other_member_index in employees_indexs:
            other_members.append(all_employees[other_member_index])        # 所有未分配的6名成员
        # 将6名未分组人员加入7个小组
        group_other_members(all_group, other_members, headmans, group_settings,
                            all_members, all_fits_groups)


def members_group1(all_employees, all_groups):
    '''
    对6名未分组人员分组
    分组后的所有分组方式个数为：7^6
    :param all_employees: 所有未分组人员
    :param all_groups: 存放所有分组情况
    :return:
    '''
    if not all_employees:
        return all_groups
    cur_member = all_employees.pop(0)        # 取出一个待分组人员
    new_all_groups = []
    for group in all_groups:
        for index, group_item in enumerate(group):
            group_item_temp = copy.deepcopy(group_item)
            group_temp = copy.deepcopy(group)
            group_item_temp.append(cur_member)                   # 将取出的待分组人员加入小组
            group_temp[index] = group_item_temp
            new_all_groups.append(group_temp)
    return members_group1(all_employees, new_all_groups)       # 返回所有分组方式


def fits_capicaty(group_index, group_all_members, group_settings, all_members):
    '''
    判断每个小组是否满足小组工作承载力
    工作承载力要求为：
        小组工作承载力 * 最低阈值比例 <= 小组人员总的工作承载力 <= 小组工作承载力 * 最好阈值比例
    :param group_index: 小组索引
    :param group_all_members: 小组人员
    :param group_settings: 分组设置
    :param all_members: 所有人员数据
    :return: 是否满足小组工作承载力要求
    '''
    capicaty_threshold = group_settings['capicaty_threshold']       # 小组工作承载力阈值比例范围
    group_capacity = group_settings['group_capacity']             # 所有小组工作承载力
    group_members = all_members[all_members['人员'].isin(group_all_members)]
    group_members['工作承载力'] = group_members['工作承载力'].astype(dtype = 'int')
    capicaty_sum = group_members['工作承载力'].sum()              # 小组人员工作承载力总和
    min_value = group_capacity[group_index] * capicaty_threshold[0]        # 小组承载力最小值
    max_value = group_capacity[group_index] * capicaty_threshold[1]        # 小组承载力最大值
    # 小组承载力需满足条件：大于小组承载力最小值，小于小组承载力最大值
    if capicaty_sum >= min_value and capicaty_sum <= max_value:
        return True
    return False


def fits_member_skill(group_all_members, group_settings, all_members):
    '''
    判断小组人员工作技能是否满足要求
    小组技能要求为：
        每个小组每项技能的人数必须大于shill_count
    :param group_all_members: 小组人员
    :param group_settings: 分组设置
    :param all_members: 所有人员数据
    :return: 是否满足小组技能要求
    '''
    skill_count = group_settings['skill_count']        # 小组每项技能需要优秀的人数
    group_members = all_members[all_members['人员'].isin(group_all_members)]     # 筛选小组人员
    flag = True
    for key, value in skill_count.items():
        skill_col = group_members[key]
        s_count = skill_col[skill_col=='优秀'].count()          # 小组每项技能优秀的人数
        if s_count < value:
            flag = False
            break
    return flag


def add_headman_justice(headman_combin, groups_combin, group_settings, all_members):
    '''
    加入小组长判断小组分配方式是否满足要求
    加入小组长后的分组个数为：
                    A7,7 * 迭代次数 * 7^6
    :param headman_combin: 小组长排列方式
    :param groups_combin: 分组方式
    :param group_settings: 分组设置
    :param all_members: 所有人员数据
    :return: 是否满足分组条件
    '''
    success_fits_flag = True
    for group_index, group_item in enumerate(groups_combin):
        group_all_members = copy.deepcopy(group_item)
        group_all_members.append(headman_combin[group_index])     # 小组长加入小组
        # 判断每个小组是否满足工作承载力要求和技能要求，只要有一个小组不满足要求终止循环
        if not fits_capicaty(group_index, group_all_members, group_settings, all_members) or \
                    not fits_member_skill(group_all_members, group_settings, all_members):
            success_fits_flag = False
            break
    return success_fits_flag


def justice_groups(headmans, group_combin, group_settings, all_members, all_fits_groups):
    '''
    判断小组分配方式是否满足分组规则主函数
    :param headmans: 小组长排列方式
    :param group_combin: 小组成员
    :param group_settings: 分组设置
    :param all_members: 所有人员数据
    :param all_fits_groups: 存储所有满足条件的分组方式
    :return:
    '''
    for headman_combin in headmans:
        # 加入小组长后的分组方式
        fits_headman_and_members = {
            'headman': headman_combin,
            'members': group_combin
        }
        # 判断每个小组是否满足分组规则
        if add_headman_justice(headman_combin, group_combin, group_settings, all_members):
            all_fits_groups.append(fits_headman_and_members)               # 存入所有满足条件的分组方式
            print('小组分配方式满足分配规则：{}'.format(fits_headman_and_members))



def show_all_fits_groups(all_fits_groups):
    '''
    展示所有满足条件的分组方式：

    第i种小组分配方式：
    小组                  组长               组员
    小组1                小组成员x            [小组成员1,小组成员2,..........]
    ...                 ...                 ...
    :param all_fits_groups: 所有满足条件的分组方式
    :return:
    '''
    for all_fits_group in all_fits_groups:
        for alloc_index, fits_group in enumerate(all_fits_group):
            print('\n第{}种小组分配方式：'.format(alloc_index))
            print('小组                  组长                 组员')
            for group_index, group_item in enumerate(fits_group['headman']):
                print(
                    '小组{}                  {}                 {}'
                        .format(group_index+1,
                                fits_group['headman'][group_index],
                                fits_group['members'][group_index])
                )



if __name__ == '__main__':
    # iter_count为迭代次数，迭代次数越高，考虑到的情况越全面，为了验证计算结果，我将该值设置为1，要列举所有情况需增大该值。
    # 保守估计：该值需大于C41,5*C36,5*C31,5*C26,5*C21,5*C16,5*C11,5.
    iter_count = 10000
    # group_settings为小组条件配置（1.每个小组的工作承载力；2.小组工作承载力阈值；3.小组技能；4.需要分配的组数）
    group_settings = {
        'group_capacity': [44504, 41023, 37620, 39580, 44308, 36903, 47050, 290998],
        'group_columns': ['capacity'],
        'capicaty_threshold': [0.9, 1.2],
        'skill_count': {
            '应知':1, '应会':1, '装表接电':1, '抄表收费':1,
            '新型业务':1, '低压设备操作':1, '高压设备操作':1,
            '设备巡视':1, '低压故障排查':1, '高压故障排查':1,
            '工程理论':1, '安全意识':1, '沟通协调能力':1
        },
        'group_n': 7
    }
    path = r'./统计表20210601.xls'                # 文件路径
    sheet_name = 'Sheet1'                           # 要读取的表格名
    # 所有人员数据列表的列名
    col_names = ['人员', '应知', '应会',
                 '装表接电', '抄表收费', '新型业务',
                 '低压设备操作', '高压设备操作', '设备巡视',
                 '低压故障排查', '高压故障排查', '工程理论',
                 '安全意识', '沟通协调能力', '折算总分', '工作承载力']
    # 从表格中读取数据，获得班长、组长、员工、所有人员的集合
    monitors, headman_array, all_employees, all_members = read_file(path, sheet_name, col_names, group_settings)
    headers = monitors['人员'].tolist()               # 选出的班长
    headmans = headman_array                         # 选出的组长
    all_employees = all_employees['人员'].tolist()          # 待分配的员工
    group_n = group_settings['group_n']                    # 分配组数

    print('---------------------------------开始分配小组-----------------------------------')
    pool = multiprocessing.Pool(processes=4)                # 进程数

    all_fits_groups = list()                   # 用于存储所有满足条件的分组
    # 启用进程开始小组分配
    pool.apply_async(members_group, (all_employees, iter_count, headmans,
                                                          group_settings, all_members, all_fits_groups,))
    pool.close()
    pool.join()

    all_fits_groups = group_deduplication(all_fits_groups)             # 对满足条件的小组分配方式去重
    number_fits_group = len(all_fits_groups)
    print(
        '\n-------------------------------------------------已完成小组分配----------------------------------------------------\n')
    print('共有{}种分配结果\n'.format(number_fits_group))
    if number_fits_group:
        print('分配结果如下：\n')
        print('班长为：\n'
              '      一班            二班\n'
              '      {}             {}\n'.format(monitors[0], monitors[1]))
        # 列出所有满足条件的小组分配方式
        show_all_fits_groups(all_fits_groups)

