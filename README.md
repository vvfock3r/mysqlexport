# mysqlexport

## 简介

导出MySQL数据到Excel

## 特性

* MySQ 流式读取数据，支持海量数据读取而不影响MySQL，默认每读取1W条休眠1秒钟进一步降低MySQL压力
* Excel 流式写入数据，但客户端内存占会随着数据而增加，大概100W条数据占用50M内存
* 支持Excel简单样式，包括但不限于行高、列宽、对齐方式等
* 支持Excel添加密码，提供基本的安全机制
* 支持Excel多工作表，默认每100W条数据会自动新建一个工作表

## 安装

```bash
# 需要Go 1.20+
go install github.com/vvfock3r/mysqlexport@latest
```

## 选项

```bash
[root@localhost ~]# ./mysqlexport --help

Export mysql to excel                                               
For details, please refer to https://github.com/vvfock3r/mysqlexport

Usage:                                                                            
  mysqlexport [flags]                                                             

General Flags:
  -v, --version                     version message
      --help                        help message
          
Log Flags:      
      --log-format string           log format (default "console")                
      --log-level string            log level (default "info")                    
      --log-output string           log output (default "stdout") 

MySQL Flags:
  -h, --host string                 mysql host (default "127.0.0.1")
  -P, --port string                 mysql port (default "3306")  
  -u, --user string                 mysql user (default "root")  
  -p, --password string             mysql password
  -d, --database string             mysql database
  -e, --execute string              execute sql command                             
      --charset string              mysql charset (default "utf8mb4")    
      --collation string            mysql collation (default "utf8mb4_general_ci")        
      --connect-timeout string      mysql connect timeout (default "5s")       
      --read-timeout string         mysql read timeout (default "30s")    
      --write-timeout string        mysql write timeout (default "30s")   
      --max-allowed-packet string   mysql max allowed packet (default "16MB")             
      --batch-size int              batch size (default 10000)
      --sleep-time string           sleep time (default "1s")
      
Excel Flags:
  -o, --output string               output xlsx file
      --excel-password string       excel-password                                
      --col-align string            col align (default "left")
      --col-width string            col-width
      --row-height string           row height
      --sheet-line int              max line per sheet (default 1000000)
      --sheet-name string           sheet name (default "Sheet")
```

## 示例

**测试环境**

* Windows 10 WPS
* Linux CentOS 7

**注意事项**

* 使用 `sz` 下载大文件会损坏，请考虑其他办法，比如SFTP、WinSCP、FTP等

**基本用法**

```bash
./mysqlexport \
	-h192.168.48.129 \
	-p"QiNqg[l.%;H>>rO9" \
	--database demo \
	-e "select * from users limit 200" \
	-o 测试.xlsx
```

**调整样式**

```bash
./mysqlexport \
	-h192.168.48.129 \
	-p"QiNqg[l.%;H>>rO9" \
	--database demo \
	-e "select * from users limit 200" \
	-o 测试.xlsx \
	--col-width="1:10,2-7:50" \
	--col-align="1-7:center" \
	--row-height="1:30,2-7:20"

# 说明
# 1、row代表行,col代表列
# 2、--col-width="1:10,2-7:40"  设置第一列宽度为10像素，第2-7列宽度为40像素
# 3、--col-align="1-7:center"   设置1-7列水平居中
# 4、--row-height="1:30,2-7:20" 设置第一行行高为30像素，第2-7行行高为20像素
# 5、单元格默认垂直居中，暂不支持自动调整
```

**其他选项**

```bash
# 设置Excel密码为123456
--excel-password=123456

# 每从MySQL中读取1W条数据程序休眠1秒，用于降低MySQL使用率，但会延长程序执行时间
--batch-size int              batch size (default 10000)
--sleep-time string           sleep time (default "1s")
```

## 截图

测试200W条数据

![image-20230413154321933](https://tuchuang-1257805459.cos.accelerate.myqcloud.com//image-20230413154321933.png)

![image-20230413155328849](https://tuchuang-1257805459.cos.accelerate.myqcloud.com//image-20230413155328849.png)

## TODO

* 代码重构
* 支持MySQL所有数据类型
* 支持Excel 更多样式调整
* 支持多工作簿
* 支持文件压缩
* 稳定性测试