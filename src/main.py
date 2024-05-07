import tkinter as tk
import tkinter.messagebox
import pandas as pd
from randSelect import *
from functools import partial

class App:
    def __init__(self, root: tk.Tk) -> None:
        self.argList = list()
        
        self.title = tk.Label(root, text="随机抽取学生", justify="left")
        self.title.config(font=("宋体", 18))
        self.title.pack()

        self.l1 = tk.Label(root, text="学生数据表(.xlsx)", justify="left")
        self.l1.config(font=("宋体", 12))
        self.l1.pack(side="top", anchor="nw")

        self.entryInputPath = tk.Entry(root, width=55)
        self.entryInputPath.pack(side="top", anchor="nw")

        self.inputPath = ""

        self.buttonGetPath = tk.Button(root, text="浏览路径..", command=self.buttonGetPathMethod)
        self.buttonGetPath.pack(side="top", anchor="nw")

        self.labelNull1 = tk.Label(root)
        self.labelNull1.pack(side="top", anchor="nw")

        self.frameNum = tk.Frame(root)
        self.frameNum.pack(side="top", anchor="nw")

        self.labelNum = tk.Label(self.frameNum, text="每批选取的人数")
        self.labelNum.config(font=("宋体", 12))
        self.labelNum.pack(side="left", anchor="nw")

        self.entryNum = tk.Entry(self.frameNum, width=4,
                                 validate="focusout",
                                 validatecommand=root.register(self.checkUpdate))
        self.entryNum.pack(side="left", anchor="nw")

        self.labelBatch = tk.Label(self.frameNum, text="选取批次数",)
        self.labelBatch.config(font=("宋体", 12))
        self.labelBatch.pack(side="left", anchor='nw')

        self.entryBatch = tk.Entry(self.frameNum, width=4,
                                   validate="focusout",
                                   validatecommand=root.register(self.checkUpdate))
        self.entryBatch.pack(side="left", anchor="nw")
        self.entryBatch.insert(0, "6")

        self.labelNull2 = tk.Label(root)  # 占位
        self.labelNull2.pack(side="top", anchor="nw")

        self.frameSelective = tk.Frame(root)
        self.frameSelective.pack(side="top", anchor="nw")

        self.labelSelective = tk.Label(self.frameSelective, text="区分本硕博学生？")
        self.labelSelective.config(font=("宋体", 12))
        self.labelSelective.pack(side="left", anchor='nw')

        self.isSeperative = tk.BooleanVar()  # 是否区分本硕博
        self.isSeperative = False
        
        self.checkbuttonSelective = tk.Checkbutton(self.frameSelective, command=self.checkbuttonSelectiveMethod)
        self.checkbuttonSelective.pack(side="left", anchor="nw")
        
        self.frameQualification = tk.Frame(root)
        self.frameQualification.pack(side="top", anchor="nw")
        
        self.modulesCheckDict = dict()
        
        self.labelBachelor = tk.Label(self.frameQualification, text="本科")
        self.labelBachelor.config(font=("宋体", 12))

        self.entryBachelor = tk.Entry(self.frameQualification, width=4, 
                                    validate="focusout", 
                                    validatecommand=root.register(partial(self.entryCheckCommand, "bachelor")))
        
        self.labelHintBachelor = tk.Label(self.frameQualification, text="      ")
        self.labelHintBachelor.config(font=("宋体", 10))
        
        self.labelMaster = tk.Label(self.frameQualification, text="硕士")
        self.labelMaster.config(font=("宋体", 12))

        self.entryMaster = tk.Entry(self.frameQualification, width=4,
                                    validate="focusout",
                                    validatecommand=root.register(partial(self.entryCheckCommand, "master")))
        # self.entryMaster.pack(side="top", anchor="nw")
        
        self.labelHintMaster = tk.Label(self.frameQualification, text="      ")
        self.labelHintMaster.config(font=("宋体", 10))
        # self.labelHintMaster.pack(side="top", anchor="nw")

        self.labelDoctor = tk.Label(self.frameQualification, text="博士")
        self.labelDoctor.config(font=("宋体", 12))

        self.entryDoctor = tk.Entry(self.frameQualification, width=4,
                                    validate="focusout",
                                    validatecommand=root.register(partial(self.entryCheckCommand, "doctor")))
        # self.entryDoctor.pack(side="top", anchor="nw")
        
        self.labelHintDoctor = tk.Label(self.frameQualification, text="      ")
        self.labelHintDoctor.config(font=("宋体", 10))
        # self.labelHintDoctor.pack(side="top", anchor="nw")

        self.frameSexRequirement = tk.Frame(root)
        self.frameSexRequirement.pack(side="top", anchor="nw")

        self.labelMale = tk.Label(self.frameSexRequirement, text="男生数")
        self.labelMale.config(font=("宋体", 12))
        self.labelMale.pack(side="left", anchor="nw")

        self.entryMale = tk.Entry(self.frameSexRequirement, width=4,
                                  validate="focusout",
                                  validatecommand=root.register(partial(self.entryCheckCommand, "male")))
        self.entryMale.pack(side="left", anchor="nw")
        
        self.labelHintMale = tk.Label(self.frameSexRequirement, text="      ")
        self.labelHintMale.config(font=("宋体", 10))
        self.labelHintMale.pack(side="left", anchor="nw")

        self.labelFemale = tk.Label(self.frameSexRequirement, text="女生数")
        self.labelFemale.config(font=("宋体", 12))
        self.labelFemale.pack(side="left", anchor="nw")

        self.entryFemale = tk.Entry(self.frameSexRequirement, width=4,
                                    validate="focusout",
                                    validatecommand=root.register(partial(self.entryCheckCommand, "female")))
        self.entryFemale.pack(side="left", anchor="nw")
        
        self.labelHintFemale = tk.Label(self.frameSexRequirement, text="      ")
        self.labelHintFemale.config(font=("宋体", 10))
        self.labelHintFemale.pack(side="left", anchor="nw")
        
        self.frameRaceRequirement = tk.Frame(root)
        self.frameRaceRequirement.pack(side="top", anchor="nw")

        self.labelHan = tk.Label(self.frameRaceRequirement, text="汉族数")
        self.labelHan.config(font=("宋体", 12))
        self.labelHan.pack(side="left", anchor="nw")

        self.entryHan = tk.Entry(self.frameRaceRequirement, width=4,
                                 validate="focusout",
                                 validatecommand=root.register(partial(self.entryCheckCommand, "han")))
        self.entryHan.pack(side="left", anchor="nw")
        
        self.labelHintHan = tk.Label(self.frameRaceRequirement, text="      ")
        self.labelHintHan.config(font=("宋体", 10))
        self.labelHintHan.pack(side="left", anchor="nw")

        self.labelMinor = tk.Label(self.frameRaceRequirement, text="少数民族数")
        self.labelMinor.config(font=("宋体", 12))
        self.labelMinor.pack(side="left", anchor="nw")

        self.entryMinor = tk.Entry(self.frameRaceRequirement, width=4,
                                   validate="focusout",
                                   validatecommand=root.register(partial(self.entryCheckCommand, "minor")))
        self.entryMinor.pack(side="left", anchor="nw")
        
        self.labelHintMinor = tk.Label(self.frameRaceRequirement, text="      ")
        self.labelHintMinor.config(font=("宋体", 10))
        self.labelHintMinor.pack(side="left", anchor="nw")

        self.modulesCheckDict = {
            "bachelor" : (self.entryBachelor, self.labelHintBachelor),
            "master" : (self.entryMaster, self.labelHintMaster),
            "doctor": (self.entryDoctor, self.labelHintDoctor),
            "male" : (self.entryMale, self.labelHintMale),
            "female" : (self.entryFemale, self.labelHintFemale),
            "han" : (self.entryHan, self.labelHintHan),
            "minor" : (self.entryMinor, self.labelHintMinor)
        }
        
        self.labelHint = tk.Label(root, text="（上四个输入框中的一个或多个可以不填写）")
        self.labelHint.config(font=("宋体", 10))
        self.labelHint.pack(side="top", anchor="nw")
        
        self.frameResult = tk.Frame(root)
        self.frameResult.pack(side="top", anchor="nw")
        
        self.resultData = None
        
        self.buttonResult = tk.Button(self.frameResult, text="显示结果", command=self.buttonResultMethod)
        self.buttonResult.pack(side="left", anchor="nw")
        
        self.buttonSave = tk.Button(self.frameResult, text="保存到路径..", command=self.bottonSaveMethod)
        self.buttonSave.pack(side="left", anchor="nw")
        
        self.textResult = tk.Text(root, wrap="char", width=55)
        self.textResult.pack(side="top", anchor="nw")
        
        self.savePath = ""
        
        self.clickTimes = 0  # 点击显示结果按钮的次数
        
        self.argList = [
            self.entryNum.get(),
            self.entryBatch.get(),
            self.entryBachelor.get(),
            self.entryMaster.get(),
            self.entryDoctor.get(),
            self.entryMale.get(),
            self.entryFemale.get(),
            self.entryHan.get(),
            self.entryMinor.get()
            ]  # 随机选择学生所需要的各种参数的list

        self.argLimits = dict()
        
        # TODO: type-in hinting label for 7 entries (bachelor, master, doctor, male, female, han, minor)
        
        # TODO: check function for entries to adjust ltype-in hinting label, and change the arg list
    
    def isUpdated(self) -> bool:
        """检查参数是否更新
        """
        return [
            self.entryNum.get(),
            self.entryBatch.get(),
            self.entryBachelor.get(),
            self.entryMaster.get(),
            self.entryDoctor.get(),
            self.entryMale.get(),
            self.entryFemale.get(),
            self.entryHan.get(),
            self.entryMinor.get()
            ] != self.argList
    
    def updateArgLists(self):
        self.argList = [
            self.entryNum.get(),
            self.entryBatch.get(),
            self.entryBachelor.get(),
            self.entryMaster.get(),
            self.entryDoctor.get(),
            self.entryMale.get(),
            self.entryFemale.get(),
            self.entryHan.get(),
            self.entryMinor.get()
            ]
    
    def checkUpdate(self, update=False):
        """检查数据是否更新，若是，更新argList，置resultData为None，删除答案显示框中的内容
        
        Args:
            update (bool): 是否更新已经显示的数据
        """
        if not self.argList:
            return
        if self.isUpdated():
            self.updateArgLists()
            self.resultData = None  # 更新结果数据，让buttonResult重新计算结果
            self.clickTimes = 0
            if update:
                self.textResult.delete("1.0", "end")  # 删除已经显示的数据
    
    def isValidInput(self, entryInput: tk.Entry, labelHint: tk.Label, checkName: str) -> bool:
        """检查entryInput中的内容是否合法，并更改labelHint中的内容

        Args:
            entryInput (tk.Entry): 输入框
            labelHint (tk.Label): 提示字符框
            checkName (str): 输入字符对应的名称(如"bachelor", "male")

        Returns:
            bool: entryInput中的参数是否合法
        """
        if not self.argLimits:
            self.argLimits = getArgLimits(self.data, batch=int(self.entryBatch.get()), num=int(self.entryNum.get()))
        if not entryInput.get():
            return False  # 防止没有内容时报错
        if int(entryInput.get()) > self.argLimits[checkName]:
            labelHint["text"] = " <={} ".format(self.argLimits[checkName])
            labelHint["fg"] = "red"  # 警告提示
            return False
        else:
            labelHint["text"] = " <={} ".format(self.argLimits[checkName])
            labelHint["fg"] = "green"  # 正常提示
            return True
    
    def entryCheckCommand(self, moduleName: str):
        """entry失去焦点时，检查数据更新，并比对此时输入的数据和输入上限的大小，如果过大，提示字体变红，否则变绿
        
        Args:
            moduleName (str): 输入的组件名称
        """
        if not self.modulesCheckDict:  # 防止在self.__init__中构建组件的时候调用该函数导致报错
            return
        entry, label = self.modulesCheckDict[moduleName]
        self.checkUpdate()  # 检查数据是否更新
        if (not self.data.empty) or self.entryInputPath.get():
            return self.isValidInput(entry, label, moduleName)
        
    def buttonGetPathMethod(self):
        import tkinter.filedialog
        self.inputPath = tkinter.filedialog.askopenfilename()
        try:
            self.data = pd.read_excel(self.inputPath)
        except FileNotFoundError:
            tkinter.messagebox.showerror(title="错误！", message="文件未找到")
            self.entryInputPath.delete(0, len(self.inputPath))
            return
        self.entryInputPath.insert(0, self.inputPath)
        
    def checkbuttonSelectiveMethod(self):
        toBeHidden = [self.labelBachelor, self.entryBachelor, self.labelHintBachelor,
                      self.labelMaster, self.entryMaster, self.labelHintMaster,
                      self.labelDoctor, self.entryDoctor, self.labelHintDoctor]
        toBeHiddenLabel = [self.labelBachelor, self.labelHintBachelor, 
                           self.labelMaster, self.labelHintMaster, 
                           self.labelDoctor, self.labelHintDoctor]
        toBeHiddenEntry = [self.entryBachelor, self.entryMaster, self.entryDoctor]
        if self.isSeperative:
            # print("hide")
            for f in toBeHiddenLabel:
                f.forget()
            for f in toBeHiddenEntry:
                f.delete(0, len(f.get()))
                f.forget()
            self.isSeperative = False
        else:
            # print("show")
            for f in toBeHidden:
                f.pack(side="left", anchor='nw')
            self.isSeperative = True

    def bottonSaveMethod(self):
        if not self.resultData:
            tkinter.messagebox.showerror(title="错误！", message="未抽取结果，请点击显示结果按钮进行抽取")
            return
        
        import tkinter.filedialog
        self.savePath = tkinter.filedialog.askdirectory()
        
        with pd.ExcelWriter(self.savePath+"/选择结果.xlsx") as writer:
            if self.isSeperative:  # 分本硕博保存
                dct = {0: "本科", 1: "硕士", 2: "博士"}
                for idx, r in enumerate(self.resultData):
                    for jdx, dfi in enumerate(r):
                        dfi.to_excel(
                            writer, 
                            sheet_name="{}第{}组".format(dct[idx], jdx+1), 
                            index=False
                        )
            else:
                for jdx, dfi in enumerate(self.resultData[0]):
                    dfi.to_excel(
                        writer, 
                        sheet_name="第{}组".format(jdx+1), 
                        index=False
                    )
        tk.messagebox.showinfo(title="完成！", message="保存成功！")
        
    def buttonResultMethod(self):
        self.checkUpdate(update=True)
        if not self.resultData:  # 没有结果，计算结果
            self.textResult.delete("1.0", "end")
            if self.data.empty:
                try:
                    df = pd.read_excel(self.inputPath)
                except FileNotFoundError:
                    tkinter.messagebox.showerror(title="错误！", message="文件未找到")
                    self.entryInputPath.delete(0, len(self.inputPath))
                    return
            else:
                df = self.data
            # 获取数据
            if self.isSeperative:
                try:
                    dfsByQual = seperateByQualificaton(df)  # 由学历区分的学生数据表
                except Exception as err:  # 数据中有未定义的培养层次
                    tk.messagebox.showerror(title="错误！", message=str(err))
                    return
            else:
                dfsByQual = [df]  # 不区分学历
            res = []
            
            if not self.isSeperative:
                try:
                    num = int(self.entryNum.get())
                except ValueError:
                    tk.messagebox.showerror(title="错误！", message="请输入每批选取的人数")
            batch = int(self.entryBatch.get())
            male = -1 if self.entryMale.get() == "" else int(self.entryMale.get())
            female = -1 if self.entryFemale.get() == "" else int(self.entryFemale.get())
            han = -1 if self.entryHan.get() == "" else int(self.entryHan.get())
            minor = -1 if self.entryMinor.get() == "" else int(self.entryMinor.get())
            for idx, dfi in enumerate(dfsByQual):
                # 按是否依据学历分类分两类，分别获取num
                if not self.isSeperative:
                    num = int(self.entryNum.get())
                    try:
                        res.append(randomSelect(dfi, num, batch, male, female, han, minor))
                    except Exception as err:
                        tk.messagebox.showerror(title="错误！", message=str(err))
                        return
                else:
                    nums = [int(self.entryBachelor.get()), int(self.entryMaster.get()), int(self.entryDoctor.get())]
                    if sum(nums) != int(self.entryNum.get()):
                        tk.messagebox.showerror(title="错误！", message="本硕博人数之和与填入的总人数不同")
                        self.entryBachelor.delete(0, len(self.entryBachelor.get()))
                        self.entryMaster.delete(0, len(self.entryMaster.get()))
                        self.entryDoctor.delete(0, len(self.entryDoctor.get()))
                        return
                    num = nums[idx]
                    try:
                        res.append(randomSelect(dfi, num, batch, male, female, han, minor))
                    except Exception as err:
                        tk.messagebox.showerror(title="错误！", message=str(err))
                        return
            self.resultData = res
            tk.messagebox.showinfo(title="成功！", message="随机抽取完成")
        
        else:  # 有结果，显示结果
            if self.clickTimes < len(self.resultData) * len(self.resultData[0]):  # resultData为list[list[Dataframe]]
                if self.clickTimes > 0:
                    self.textResult.delete("1.0", "end")
                txt = ""
                df = self.resultData[self.clickTimes // len(self.resultData[0])][self.clickTimes % len(self.resultData[0])]
                for line in df.iloc:
                    for ele in line:
                        txt += f"{ele} "
                    txt += "\n"
                self.textResult.insert("1.0", txt)
                self.clickTimes += 1
            else:
                self.clickTimes = 0
                tk.messagebox.showinfo(title="完成！", message="数据显示完成，再次点击按钮重新开始显示")
                self.textResult.delete("1.0", "end")


if __name__ == "__main__":
    root = tk.Tk()
    root.title("随机抽取学生")
    root.geometry("400x500")
    app=App(root)
    root.mainloop()