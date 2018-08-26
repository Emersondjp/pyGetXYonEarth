
import xlrd, xlwt

fileName=u'长春市2000年部分汽车企业概况统计';

# 公司名称所在行数列表
corpStartRow=[1, 31, 60, 87,115];
# 公司信息结束行数列表
corpEndRow = [30, 59, 86, 114, 141];

# 输入文件名称
inFile = u'D:\\Data\\input\\'+fileName+u'.xls';
# 输出文件名称
outFile = u'D:\\Data\\output\\'+fileName+u'.xls';

out = xlwt.Workbook();
book = xlrd.open_workbook(inFile);
# sheet名称
outSheet = out.add_sheet(fileName);


################ 下面的部分不要修改 ###################
outRow0 = [u'企业名称',u'通讯地址',u'邮政编码',u'电话',u'主要产品名称型号及产量',\
           u'①全年工业总产值(1990年不变价),万元',u'②全年工业总产值(现行价),万元'\
           u'③汽车工业产值(1990年不变价),万元',u'④工业增加值(现行价),万元',\
           u'⑤年末从业人数合计,人',u'⑥年末资产总计,万元',u'⑦产品销售收入,万元'];
for i in range(0,len(outRow0)):
    outSheet.write(0,i,outRow0[i]);

table = book.sheet_by_index(0);

# 表格列数值
ncols = table.ncols;

wRow=1;

for rowIndex in range(0, len(corpStartRow)):
    corpRowList = range(corpStartRow[rowIndex], corpEndRow[rowIndex]+1);
    for col in range(1, ncols):
        corpName = str(table.cell(corpRowList[0], col).value);
        if(corpName == ''): continue ;
        outSheet.write(wRow, 0, corpName);
        print(corpName);


        corpAddr = ''.join([ str(table.cell(x,col).value) for x in corpRowList[1:4]])
        outSheet.write(wRow, 1, corpAddr);

        corpPostCode = str(table.cell(corpRowList[5],col).value);
        outSheet.write(wRow, 2, corpPostCode.replace('.0',''));
            
        corpTele = str(table.cell(corpRowList[6],col).value)+'-'+str(table.cell(corpRowList[7],col).value);
        outSheet.write(wRow, 3, corpTele.replace('.0',''));

        corpMainProduct = ''.join([str(table.cell(x,col).value) for x in corpRowList[8:-8]]);
        outSheet.write(wRow, 4, corpMainProduct);

        corpIndustryValue = str(table.cell(corpRowList[-7],col).value);
        outSheet.write(wRow, 5, corpIndustryValue.replace('.0',''));
            
        corpIndustryValue2 = str(table.cell(corpRowList[-6],col).value);
        outSheet.write(wRow, 6, corpIndustryValue2.replace('.0',''));

        corpMobileIndustryValue = str(table.cell(corpRowList[-5],col).value);
        outSheet.write(wRow, 7, corpMobileIndustryValue.replace('.0',''));
            
        corpIndustryIncreaseValue = str(table.cell(corpRowList[-4],col).value);
        outSheet.write(wRow, 8, corpIndustryIncreaseValue.replace('.0',''));
           
        corpWorkerNo = str(table.cell(corpRowList[-3],col).value);
        outSheet.write(wRow, 9, corpWorkerNo.replace('.0',''));
            
        corpAsset = str(table.cell(corpRowList[-2],col).value);
        outSheet.write(wRow, 10, corpAsset.replace('.0',''));
            
        corpInCome = str(table.cell(corpRowList[-1], col).value);
        outSheet.write(wRow, 11, corpInCome.replace('.0',''));
        wRow = wRow + 1;

out.save(outFile);
