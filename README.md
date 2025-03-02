# KlbqAutoPiano
 卡拉比丘自动弹琴
 该软件主要是为了能够实现卡拉彼丘琴房自动化演奏。具体流程是先捕捉游戏窗口，然后通过点按射击琴房的左上和右下角计算4x4按键的坐标，然后加载乐谱，将鼠标移动到游戏内，按快捷键自动演奏。鼠标移动功能比较基础。
 可能会被反作弊系统识别，不要用大号使用！！！
 后续需要改进鼠标移动方式。目前还有一些比较摸不清头脑的bug，如有问题请提交讨论。
![QQ截图20250302232525](https://github.com/user-attachments/assets/9fb6b779-1916-4910-86f7-9bfe365cb925)

[乐谱格式test.json](https://github.com/user-attachments/files/19043674/test.json){
    "bpm": 60,
    "notes": [
        {"beat": 0.5, "block": 1},
        {"beat": 0.5, "block": 3},
        {"beat": 1.0, "block": 5},
        {"beat": 0.5, "block": 7},
        {"beat": 0.5, "block": 9},
        {"beat": 0.5, "block": 11},
        {"beat": 2.0, "block": 13},
        {"beat": 1.0, "block": 15},
        {"beat": 0.5, "block": 16},
    ]
}
