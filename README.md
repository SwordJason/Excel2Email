# Excel2Email
Send salary or other sensitive information directly to employee themself with one script.


目前通用解决方案是通过微软的Office套件， 通过excel + word + outlook去发送邮件， 对于只是发送一个工资条等需求来说， 有点浪费了， 所以开发了这个基于python的脚本。

How to configuration:
1. config.json里面
    1. 邮件信息： mail 有
        a. smtp服务器地址， 
        b. email表示邮箱地址， 
        c. password邮箱密码/Token, 
        d. subject邮件主题， 
        e. toEmail收件人的邮件地址在excel中对应的列名
        f. log对应的是每条邮件发送的时候， 输出到命令行对应的excel的列名（为了让你看清楚当前发送的是那一个人/ID)
    2. excel对应的是excel的地址（相对地址，或者绝对地址均可）
    3. template.html表示的是发送的内容模板（html形式)
2. template.html是要发送的内容模板， 其中${}表示转义， 通过在${Excel的列名}，脚本会自动获取对应的信息
3. excel里面要求首行是标题行，

How to use:
1. 将项目打包下载到本地，解压后等待使用.
2. 安装Python(已经安装过的请忽略），访问： https://www.python.org/downloads/ （一个Bug，如果选择给所有用户安装， 会导致后续的pip install都失败)
3. pip install pandas。 （项目通过pandas读取的excel内容）
4. python Excel2Email.py就可以运行.

