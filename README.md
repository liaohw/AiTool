# 报销系统生成模版，用于批量导入和小票打印

## execl说明：

* N列为需要调整的配置；周六日与休息日，不同金额区分
* 在“年+月+星期”基础上，支持按Sheet2日期过滤；“个性化日期”参考Sheet3按工作日金额*人数计算总金额生成小票；
* 在“星期都不选”情况下，直接以Sheet3生成；
* Sheet3支持30、60、大额填写
* doc文件不存在,会新建但格式是宽屏，建议用模版doc


## 使用说明：

打开“小票生成.xlsm”，修改配置，单击“执行生成”；生成execl和word

## 结果如下：

![excel.png](https://github.com/liaohw/AiTool/blob/master/res/excel.png)
![word.png](https://github.com/liaohw/AiTool/blob/master/res/word.png)




