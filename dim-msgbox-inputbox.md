# VBScript基础——dim、msgbox和inputbox函数

## VBS 简介

VBScript是 Visual Basic Script 的简称，即 Visual Basic 脚本语言，有时也被缩写为 VBS。

VBScript 是微软开发的一种脚本语言。使用 VBScript，可通过 Windows 脚本宿主调用 COM，所以可以使用 Windows 操作系统中可被使用的程序库。

VBScript 一般被用在以下个方面：VBScript 经常被用来完成重复性的Windows 操作系统任务；用来指挥客户方的网页浏览器。在这一方面，VBS 与JavaScript 是竞争者，因为本文的实验环境基于 Window 平台，为了达到更好的兼容性和性能选用 VBScript。

由于VBScript可以通过[Windows脚本宿主](https://baike.baidu.com/item/Windows脚本宿主?fromModule=lemma_inlink)调用[COM](https://baike.baidu.com/item/COM?fromModule=lemma_inlink)，因而可以使用Windows[操作系统](https://baike.baidu.com/item/操作系统?fromModule=lemma_inlink)中可以被使用的[程序库](https://baike.baidu.com/item/程序库?fromModule=lemma_inlink)，比如它可以使用Microsoft Office的库，尤其是使用Microsoft [Access](https://baike.baidu.com/item/Access?fromModule=lemma_inlink)和[Microsoft](https://baike.baidu.com/item/Microsoft?fromModule=lemma_inlink)SQL Server的程序库，当然它也可以使用其它程序和操作系统本身的库。

VBScript是基于Visual Basic程序语言的脚本语言,是IIS(互联网信息服务,InternetInformation Services)的默认源程序语言。VBScript最开始是通过事件驱动来扩展客户端HTML的功能,可在网页上处理、控制对象,它能与HTML页面很好的结合使用,VBScript可是操作HTML页面,还可对页面中的事件做出响应。另外,VBScript还提供了一些应用对象,使编写者更方便地编写脚本,用于实现一些特有功能。

## Dim函数

dim函数的用法

`dim 变量名`

1.变量名
这个字符串可以是任何英文字符，这个函数可以创建这个变量
例如：创建名为a的变量
`dim a`

## Msgbox函数

msgbox函数的用法
`msgbox"内容",格式,"标题"`

1.内容
这个字符串值可以是任何字符，这个字符串会在弹窗提示的主要部分显示。
例如：显示内容为“Microsoft VBS”
`msgbox"Microsoft VBS"`

2.格式
这个数值可以是下面的任何数字相加的结果：

| 值   | 功能                                                         |
| ---- | ------------------------------------------------------------ |
| 0    | 只显示确定按钮（默认）。                                     |
| 1    | 显示确定和取消按钮。                                         |
| 2    | 显示放弃、重试和忽略按钮。                                   |
| 3    | 显示是、否和取消按钮。                                       |
| 4    | 显示是和否按钮。                                             |
| 5    | 显示重试和取消按钮。                                         |
| 16   | 显示临界信息图标。                                           |
| 32   | 显示警告查询图标。                                           |
| 48   | 显示警告消息图标。                                           |
| 64   | 显示信息消息图标。                                           |
| 0    | 第一个按钮为默认按钮。                                       |
| 256  | 第二个按钮为默认按钮。                                       |
| 512  | 第三个按钮为默认按钮。                                       |
| 768  | 第四个按钮为默认按钮。                                       |
| 0    | 应用程序模式：用户必须响应消息框才能继续在当前应用程序中工作。 |
| 4096 | 系统模式：在用户响应消息框前，所有应用程序都被挂起。         |

例如：显示"ABC"并显示信息图标
`msgbox"ABC",64,""`

3.标题
这个字符串值可以是任何字符，这个字符串会被显示在提示窗口的标题上。
例如：显示内容为"ABC"并显示信息图标并显示标题"CBA"
`msgbox"ABC",64,"CBA"`

4.返回值
将"msgbox"函数赋值给变量时，赋的值可以参考下表：

| 值   | 被点击的按钮 |
| ---- | ------------ |
| 1    | 确定         |
| 2    | 取消         |
| 3    | 放弃         |
| 4    | 重试         |
| 5    | 忽略         |
| 6    | 是           |
| 7    | 否           |

## Inputbox函数

inputbox函数的用法：
`inputbox"问题","标题","默认填充"`

1.问题
这个字符串值可以是任何字符，这个字符串会被显示在问题上
例如：显示问题是“Are you okay?”
`inputbox"Are you okay?"`

2.标题
这个字符串值可以是的任何字符，这个字符串会被显示在问题的标题上
例如：显示问题是"Yes?"且标题为"Liston"
`inputbox"Yes?","Liston"`

3.默认填充
这个字符串值可以是的任何字符，这个字符串会自动填充至回答
例如：显示问题是"Yes?"且标题为"Liston"并自动填充“What?”
`inputbox"Yes?","Liston","What?"`
