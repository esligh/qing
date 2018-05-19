# qing

### 准备工作：

将下载下来的各家资源先放在backup目录下备份，然后再对文档进行按照各家资源重命名，
比如晶茂的资源文档命名为晶茂.xlsx，将命名后的文档放在doc目录下

### 运行程序：

- 点击run.bat运行程序，第一次运行程序可能比较慢，需要解析excel文档内容，并生成对应的
pk缓存文件，比如万达的资源会生成wanda.pk文件，月票房的资源会生成temp.pk，
如果更改了某个资源文件(比如增加或者减少了某些记录)需要手动删除对应的pk文件，然后
重新运行程序。

- 程序运行完成后会写月票房的excel文档，针对每一行做资源的归属，如果某一个资源
在月票房中不存在就将其添加到末尾。


__注：每次运行前需要关掉doc下所有打开的文档，在程序运行过程中不要打开excel文件。__
