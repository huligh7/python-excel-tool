# python-excel-tool
python tool for office
# Excel表格文件小工具
全程搜索引擎指导下写的小工具。
## 合并Excel文件
   仅合并表结构相同的文件
## Excel数据匹配
   匹配两个表格的数据，    
   打开需要匹配的两个表格文件    
   根据表头选择两列数据，    
   有相同的值会被匹配出来    
   可以选择一个其他列的数据    
   文件A   

![image](https://user-images.githubusercontent.com/121083401/208793977-91a499d8-9c07-4103-83f7-1403a83b5686.png)


   文件B

![image](https://user-images.githubusercontent.com/121083401/208794263-ca20ba1a-3042-412f-aba1-29093f9ee06a.png)

   打开两个文件   
     选择两个需要匹配的列比如上面的A文件的c列和B文件的go列   
     得到：   
     3   
     8   
     选择A文件同时选择了另外一列，比如a   
     结果如下   
     3  1   
     8  7   
     选择B文件同时选择了另外一列，比如ee   
     结果如下   
     3  3   
     8  6   
     选择A文件同时选择了另外一列，比如a，选择B文件同时选择了另外一列，比如ee，   
     结果如下   
     3  1  3   
     8  7  6   
