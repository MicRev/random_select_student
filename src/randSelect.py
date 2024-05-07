import pandas as pd
import random
import numpy as np

def _solve(b: np.array, maxNum: list) -> np.array:
    """解方程Ax=b，x为男汉族，男少数民族，女汉族，女少数民族人数。由于该方程是个不定方程，所以遍历女少数民族数直到有合理解。

    Args:
        b (np.array): 男生数，女生数，汉族数，少数民族数
        maxNum (list): 解中每一项最大允许的值

    Returns:
        np.array: 男汉族，男少数民族，女汉族，女少数民族人数
    """
    print(b)
    for d in range(int(max(b))):
        x = np.array([b[0] - b[3] + d, b[3] - d, b[1] - d ,d])
        if all([xi >= 0 for xi in x]) and all([x[i] <= maxNum[i] for i in range(4)]):
            break
    else:
        raise Exception("solution not found")
    """
              汉族(b[2])   少数民族(b[3])
    男生(b[0]) b[0]-b[3]+d b[3] - d
    女生(b[1]) b[1]-d      d
    """
    return x


def seperateByQualificaton(data: pd.DataFrame) -> list[pd.DataFrame]:
    """把学生数据表按学历分为三个表

    Args:
        data (pd.DataFrame): 学生数据表

    Returns:
        list[pd.DataFrame]: 本科生、硕士生、博士生数据表的列表
    """
    _, _, idx = findIndexes(data)
    tab_bcl = data[data[idx] == "本科"]
    tab_mst = data[data[idx] == "硕士"]
    tab_dct = data[data[idx] == "博士"]
    return [tab_bcl, tab_mst, tab_dct]


def randomSelect(
        data: pd.DataFrame, num: int, batch=6, male=-1, female=-1, han=-1, minor=-1
    ) -> list[pd.DataFrame]:
    """随机选择几批学生，每批不重复
    
    把学生数据按性别、民族分为四个子表，若指定了不同性别、民族需要选择的人数，则按人数选取；若无，则随机分配。对四个子表分别按需要的人数随机选取学生，并合并为一个表返回。

    Args:
        data (pd.DataFrame): 学生数据的DataFrame
        num (int): 每批选择的学生数量
        batch (int): 选择的批数. Defaults to 6.
        male (int): 男同学数. Defaults to -1.
        female (int): 女同学数. Defaults to -1.
        han (int): 汉族同学数. Defaults to -1.
        minor (int): 少数民族同学数. Defaults to -1.

    Returns:
        list(pd.DataFrame): 长度为batch的列表，元素为被选择的同学DataFrame
    """
    print(num, batch, male, female, han, minor)
    argLimits = getArgLimits(data, batch=6, num=num)
    print(argLimits)
    if num > argLimits["num"]:  # 选择的学生数过多
        raise Exception("选择的总人数超过学生人数总和. 最多选择{}个".format(argLimits["num"]))
    
    indexSexualRaceQual = findIndexes(data)
    sex, race, _ = indexSexualRaceQual
    
    # 分为男生表和女生表
    if sex != -1: # 有性别列
        df_male = data[data[sex] == "男"]
        df_female = data[data[sex] == "女"]
    else:
        df_male = data.sample(frac=0.5)
        df_female = data.drop(labels=df_male.index)
    dfs = [df_male, df_female]
    _dfs = []
    
    # 进一步分为汉族表和少数民族表，每个表（性别）分别分为两个
    for df in dfs:
        if race != -1: # 有民族列
            df_han = df[df[race] == "汉族"]
            df_minor = df[df[race] != "汉族"]
        else: # 无民族列
            df_han = df.sample(frac=0.5)
            df_minor = df.drop(labels=df_han.index)
        _dfs.append(df_han)
        _dfs.append(df_minor)
    dfs = _dfs  # 分为四个dataframe，分别是男汉族，男少数民族，女汉族，女少数民族
    
    batch_dfs = []  # 每批选取的学生表
    
    trialCounts = 0  # 已经尝试的次数
    while len(batch_dfs) < batch:
        # 随机选取
        # 计算四个dataframe应该选取的人数，若缺省则随机，再在每个dataframe里随机分配
        
        if trialCounts >= 100:  # 防止因为未知原因导致的无解
            raise Exception("参数不合法，请检查后重新输入")
        
        b = np.zeros(4)  # 男生数，女生数，汉族数，少数民族数
        if male >= 0 and female >= 0:
            if male + female != num:
                raise Exception("输入的男生女生人数总和与每批选择的人数不同")
            b[0] = male
            b[1] = female
        elif male >=0 or female >= 0:
            if male >= 0:
                female = num - male
            else: male = num - female
            print(male, female)
            if not (male <= argLimits["male"] and female <= argLimits["female"]):
                exceptionText = "输入的人数过多。最大数量: "
                if male <= argLimits["male"]:
                    exceptionText += "男生{} ".format(argLimits["male"])
                if female <= argLimits["female"]:
                    exceptionText += "女生{}".format(argLimits["female"])    
                raise Exception(exceptionText)
        else:  # Default
            b[0] = ( _ := random.randint(0, num))
            b[1] = num - _
        
        if han >= 0 and minor >= 0:
            if han + minor != num:
                raise Exception("选择的不同民族人数总和与每批选择的人数不同")
            b[2] = han
            b[3] = minor
        elif han >=0 or minor >= 0:
            if han >= 0:
                minor = num - han
            else: han = num - minor
            if not(han <= argLimits["han"] and minor <= argLimits["minor"]):
                exceptionText = "输入的人数过多。最大数量: "
                if han < argLimits["han"]:
                    exceptionText += "汉族{} ".format(argLimits["han"])
                if minor < argLimits["minor"]:
                    exceptionText += "少数民族{}".format(argLimits["minor"])    
                raise Exception(exceptionText)
        else:  # Default
            b[2] = ( _ := random.randint(0, num))
            b[3] = num - _
        # print(b)
        try:
            x = _solve(b, [df.shape[0] for df in dfs])  # 获取四个子表需要选取的学生数, 第二个参数为男汉族、男少数民族、女汉族、女少数民族没被选过的总人数
        except Exception:  # Unfortunately，没有合法解，再试一次(
            trialCounts += 1
            continue
        # print(x)
        # 对四个表随机选取学生
        res =[]
        for idx, (df, n) in enumerate(zip(dfs, x)):
            n = int(n)
            indexes = random.sample(list(df.index), n)  # 选取学生的索引列
            res += [df.loc[i] for i in indexes]
            indexes = pd.Index(indexes)
            dfs[idx].drop(indexes, inplace=True)  # 删除已经选择的学生，防止重复
            # print(dfs[idx].shape[0])
        # print("--------------")
        df_res = pd.DataFrame(data=res)
        batch_dfs.append(df_res)
    
    return batch_dfs

def findIndexes(data: pd.DataFrame) -> list:
    """找到数据表中性别、民族、学历所在的列

    Args:
        data (pd.DataFrame): 数据表

    Returns:
        list: 性别、民族、学历所在的列的索引（列名）
    """
    res_index = [-1, -1, -1]
    for idx, unit in enumerate(data.iloc[1]):
        if type(unit) == type(str()):
            if unit == "男" or unit == "女":
                res_index[0] = idx
            if unit[-1] == "族":
                res_index[1] = idx
            if unit == "本科" or unit == "硕士" or unit == "博士":
                res_index[2] = idx
    res = [data.columns[idx] if idx != -1 else -1 for idx in res_index]
    return res

def getArgLimits(data: pd.DataFrame, batch, num) -> dict:
    """获取data中参数的相应上限，返回字典

    Args:
        data (pd.DataFrame): 学生数据表.
        batch (int): 选的批数.
        num (int): 学生总人数.

    Returns:
        dict: 对应元素在数据表中在该批数下最多取多少个
    """
    res = dict()
    
    # 检查输入总人数是否合法
    res["num"] = data.shape[0] // batch
    
    idxSex, idxRace, idxQualification = findIndexes(data)
    
    # print(idxQualification, idxSex, idxRace)
    if idxQualification != -1:
        tab_bcl = data[data[idxQualification] == "本科"]
        tab_mst = data[data[idxQualification] == "硕士"]
        tab_dct = data[data[idxQualification] == "博士"]
        res["bachelor"] = min(tab_bcl.shape[0] // batch, num)
        res["master"] = min(tab_mst.shape[0] // batch, num)
        res["doctor"] = min(tab_dct.shape[0] // batch, num)
    else:
        res["bachelor"] = num
        res["master"] = num
        res["doctor"] = num
    
    if idxSex != -1:
        tab_male = data[data[idxSex] == "男"]
        tab_female = data[data[idxSex] == "女"]
        res["male"] = min(tab_male.shape[0] // batch, num)
        res["female"] = min(tab_female.shape[0] // batch, num)
    else:
        res["male"] = num
        res["female"] = num
    
    if idxRace != -1:
        tab_han = data[data[idxRace] == "汉族"]
        tab_minor = data[data[idxRace] != "汉族"]
        res["han"] = min(tab_han.shape[0] // batch, num)
        res["minor"] = min(tab_minor.shape[0] // batch, num)
    else:
        res["han"] = num
        res["minor"] = num
    
    return res


if __name__ == "__main__":
    df = pd.read_excel("练习名单.xlsx")
    print(randomSelect(df, batch=6, num=10, male=7, female=3, han=6, minor=4))