<!--
 * @@Descripttion: 
 * @@version: 
 * @@encoding: utf-8
 * @@Author: qiurongcan
 * @Date: 2025-09-19 16:13:14
 * @LastEditTime: 2025-09-22 11:32:40
-->
# 2025使用python生成word版本简历

## 文件说明
```shell
|-简历编码目录.xlsx             待生成的简历文件数量
|-简历投递变量.docx             所有原始变量存储的文件
|-男-计算机专业.docx            demo
```
## 代码说明
`information.json`用于存储生成简历的变量  
`gen_resume_docx2025.py` 生成经历的代码 适用于Panel A  

**run**
```shell
# panelA
python gen_resume_docx2025.py # 民族
```
**run others**
```shell

# 生成PanelB
python gen_resume_panelB.py # 海归


# 生成PanelC
python gen_resume_panelC.py # 专业

# 生成PanelD
python gen_resume_panelD.py # 第一学历

# 生成PanelE
python gen_resum_panelE.py # 大专生
```

## 注意事项
1. 读取excel表格时候,可以修改代码读取不同的sheet
```python

sheet = 'PanelA_民族' # choices = [PanelA_民族, PanelB_海归, ...] 详细根据 简历编码目录.xlsx文件
df = pd.read_excel(r'简历编码目录.xlsx', index_col=None, sheet_name=sheet)
```
2. 详细了解一下 `information.json` 文件，如果简历出现了问题，则修改information中的变量，实现BUG的修改



