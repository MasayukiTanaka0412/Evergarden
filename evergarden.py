import logging
import os
import pandas as pd

logging.basicConfig(level=logging.INFO)
logging.info('App Start')

templateNameJ = "newsletter_j.html"
templateNameE = "newsletter_e.html"
contentsFile = "Contents.xlsx"
path = os.getcwd()

logging.info("Parameters")
logging.info("templateNameJ:{}".format(templateNameJ))
logging.info("templateNameE:{}".format(templateNameE))
logging.info("contentsFile:{}".format(contentsFile))
logging.info("path: {}".format(path))


templateJ = None
templateE = None

with open(templateNameJ,encoding="UTF8") as f:
    templateJ = f.read()
    logging.info("================================================================")
    logging.info("Japanese template")
    logging.info(templateJ)
    logging.info("================================================================")

with open(templateNameE,encoding="UTF8") as f:
    templateE = f.read()
    logging.info("================================================================")
    logging.info("English template")
    logging.info(templateE)
    logging.info("================================================================")

df = pd.read_excel(os.path.join(path,contentsFile), sheet_name='ContentsJ')
logging.info(df)
for index, row in df.iterrows():
    outputJ = templateJ
    for indexName in row.index:
        outputJ = outputJ.replace(indexName,row[indexName])

df = pd.read_excel(os.path.join(path,contentsFile), sheet_name='ContentsE')
logging.info(df)
for index, row in df.iterrows():
    outputE = templateE
    for indexName in row.index:
        outputE = outputE.replace(indexName,row[indexName])

outputdir ="output"
if not os.path.isdir(outputdir):
    os.mkdir(outputdir)

with open(os.path.join(outputdir,templateNameJ),encoding="UTF8",mode="w") as f:
    f.write(outputJ)

with open(os.path.join(outputdir,templateNameE),encoding="UTF8",mode="w") as f:
    f.write(outputE)
logging.info('App End')