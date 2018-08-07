VERSION 5.00
Begin VB.Form form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "斗败俄罗斯方块1.11"
   ClientHeight    =   11430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14370
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":076A
   ScaleHeight     =   11430
   ScaleWidth      =   14370
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer 奖励时间系统 
      Left            =   5280
      Top             =   8760
   End
   Begin VB.Timer 消失时间系统 
      Left            =   5280
      Top             =   7920
   End
   Begin VB.Timer 精力十足时间系统 
      Left            =   5280
      Top             =   7080
   End
   Begin VB.Timer 虎不发威时间系统 
      Left            =   5280
      Top             =   6360
   End
   Begin VB.Timer 树木皇后时间系统 
      Left            =   5160
      Top             =   5640
   End
   Begin VB.Timer 材料小偷时间系统 
      Left            =   5160
      Top             =   5040
   End
   Begin VB.Timer 左右时间系统 
      Left            =   4920
      Top             =   3840
   End
   Begin VB.Timer 能量小偷时间系统 
      Left            =   5040
      Top             =   3600
   End
   Begin VB.Timer 紧张时间系统 
      Left            =   5040
      Top             =   2640
   End
   Begin VB.Timer 游戏变脸时间系统 
      Left            =   6600
      Top             =   5160
   End
   Begin VB.Timer 开始按钮动画时间系统 
      Left            =   7920
      Top             =   8280
   End
   Begin VB.Timer 流星时间系统 
      Left            =   7920
      Top             =   7560
   End
   Begin VB.Timer 桌面树时间系统 
      Left            =   7800
      Top             =   6840
   End
   Begin VB.Timer 桌面粒子时间系统 
      Left            =   7800
      Top             =   6120
   End
   Begin VB.Timer 桌面星时间系统 
      Left            =   7800
      Top             =   5280
   End
   Begin VB.Timer timer1 
      Interval        =   1000
      Left            =   7800
      Top             =   2640
   End
   Begin VB.Label 提示不同能力 
      BackStyle       =   0  'Transparent
      Caption         =   "不同的敌人有不同的能力-->"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   63
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   62
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   61
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   60
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   59
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   58
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   57
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   56
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   55
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   54
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   53
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   52
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   51
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   50
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   49
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   48
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   47
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   46
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   45
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   44
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   43
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   42
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   41
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   40
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   39
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   38
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   37
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   36
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   35
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   34
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   33
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   32
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   31
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   30
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   29
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   28
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   27
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   26
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   25
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   24
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   23
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   22
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   21
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   20
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   19
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   18
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   17
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   16
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   15
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   14
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   13
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   12
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   11
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   10
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   9
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   8
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   7
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   6
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   5
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   4
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   3
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   2
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image 奖励 
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image 买精力十足 
      Height          =   1200
      Left            =   6120
      Picture         =   "Form1.frx":AE8D
      ToolTipText     =   "80"
      Top             =   7800
      Width           =   1200
   End
   Begin VB.Image 买学猴王 
      Height          =   1200
      Left            =   6000
      Picture         =   "Form1.frx":FD97
      ToolTipText     =   "70"
      Top             =   5880
      Width           =   1200
   End
   Begin VB.Image 买虎不发威 
      Height          =   1200
      Left            =   6000
      Picture         =   "Form1.frx":14CA1
      ToolTipText     =   "60"
      Top             =   4200
      Width           =   1200
   End
   Begin VB.Image 买招财猫 
      Height          =   1200
      Left            =   6000
      Picture         =   "Form1.frx":19BAB
      ToolTipText     =   "50"
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Image 商店 
      Height          =   10530
      Left            =   8040
      Picture         =   "Form1.frx":1EAB5
      Top             =   960
      Width           =   8190
   End
   Begin VB.Image 升级界面 
      Height          =   23445
      Left            =   8400
      Picture         =   "Form1.frx":3034D
      Stretch         =   -1  'True
      Top             =   480
      Width           =   17040
   End
   Begin VB.Image 结果 
      Height          =   1500
      Index           =   1
      Left            =   6000
      Picture         =   "Form1.frx":56327
      Top             =   600
      Width           =   1500
   End
   Begin VB.Image 结果 
      Height          =   1500
      Index           =   0
      Left            =   6240
      Picture         =   "Form1.frx":5DEE1
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label 能量数字 
      Caption         =   "Label1"
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label 游戏时间数字 
      Caption         =   "Label1"
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label 电脑名 
      Caption         =   "Label1"
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label 玩家名 
      Caption         =   "Label1"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image e键 
      Height          =   495
      Left            =   6480
      Top             =   840
      Width           =   615
   End
   Begin VB.Image w键 
      Height          =   375
      Left            =   6480
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image q键 
      Height          =   495
      Left            =   6360
      Top             =   960
      Width           =   975
   End
   Begin VB.Image 能量 
      Height          =   495
      Left            =   6960
      Top             =   960
      Width           =   735
   End
   Begin VB.Image 游戏时间 
      Height          =   735
      Left            =   6360
      Top             =   840
      Width           =   735
   End
   Begin VB.Image 电脑等级 
      Height          =   255
      Left            =   6480
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image 玩家等级 
      Height          =   375
      Left            =   6240
      Top             =   1200
      Width           =   615
   End
   Begin VB.Image 电脑二 
      Height          =   495
      Left            =   6360
      Top             =   960
      Width           =   855
   End
   Begin VB.Image 电脑一 
      Height          =   615
      Left            =   5880
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image 玩家二 
      Height          =   495
      Left            =   6360
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image 玩家一 
      Height          =   495
      Left            =   6600
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   63
      Left            =   6360
      Top             =   720
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   62
      Left            =   6480
      Top             =   960
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   61
      Left            =   6360
      Top             =   720
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   60
      Left            =   6480
      Top             =   960
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   59
      Left            =   6360
      Top             =   720
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   58
      Left            =   6480
      Top             =   960
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   57
      Left            =   6360
      Top             =   720
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   56
      Left            =   6480
      Top             =   960
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   55
      Left            =   6360
      Top             =   720
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   54
      Left            =   6480
      Top             =   960
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   53
      Left            =   6360
      Top             =   720
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   52
      Left            =   6480
      Top             =   960
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   51
      Left            =   6360
      Top             =   720
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   50
      Left            =   6480
      Top             =   960
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   49
      Left            =   6360
      Top             =   720
      Width           =   735
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   48
      Left            =   6480
      Picture         =   "Form1.frx":65A9B
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   47
      Left            =   6360
      Picture         =   "Form1.frx":66B35
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   46
      Left            =   6480
      Picture         =   "Form1.frx":67BCF
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   45
      Left            =   6360
      Picture         =   "Form1.frx":68C69
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   44
      Left            =   6480
      Picture         =   "Form1.frx":69D03
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   43
      Left            =   6360
      Picture         =   "Form1.frx":6AD9D
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   42
      Left            =   6480
      Picture         =   "Form1.frx":6BE37
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   41
      Left            =   6360
      Picture         =   "Form1.frx":6CED1
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   40
      Left            =   6480
      Picture         =   "Form1.frx":6DF6B
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   39
      Left            =   6360
      Picture         =   "Form1.frx":6F005
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   38
      Left            =   6480
      Picture         =   "Form1.frx":7009F
      Top             =   960
      Width           =   360
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   37
      Left            =   6360
      Picture         =   "Form1.frx":70809
      Top             =   720
      Width           =   360
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   36
      Left            =   6480
      Picture         =   "Form1.frx":70F73
      Top             =   960
      Width           =   360
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   35
      Left            =   6360
      Picture         =   "Form1.frx":716DD
      Top             =   720
      Width           =   360
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   34
      Left            =   6480
      Picture         =   "Form1.frx":71E47
      Top             =   960
      Width           =   360
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   33
      Left            =   6360
      Picture         =   "Form1.frx":725B1
      Top             =   720
      Width           =   360
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   32
      Left            =   6480
      Picture         =   "Form1.frx":72D1B
      Top             =   960
      Width           =   360
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   31
      Left            =   6360
      Picture         =   "Form1.frx":73485
      Top             =   720
      Width           =   360
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   30
      Left            =   6480
      Picture         =   "Form1.frx":73BEF
      Top             =   960
      Width           =   1050
   End
   Begin VB.Image 游戏图库 
      Height          =   360
      Index           =   29
      Left            =   6360
      Picture         =   "Form1.frx":75139
      Top             =   720
      Width           =   1050
   End
   Begin VB.Image 游戏图库 
      Height          =   180
      Index           =   28
      Left            =   6480
      Picture         =   "Form1.frx":76683
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   180
      Index           =   27
      Left            =   6360
      Picture         =   "Form1.frx":76C3D
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   180
      Index           =   26
      Left            =   6480
      Picture         =   "Form1.frx":771F7
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   180
      Index           =   25
      Left            =   6360
      Picture         =   "Form1.frx":777B1
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   24
      Left            =   6480
      Picture         =   "Form1.frx":77D6B
      Top             =   960
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   23
      Left            =   6360
      Picture         =   "Form1.frx":7CC75
      Top             =   720
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   22
      Left            =   6480
      Picture         =   "Form1.frx":81B7F
      Top             =   960
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   21
      Left            =   6360
      Picture         =   "Form1.frx":86A89
      Top             =   720
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   20
      Left            =   6480
      Picture         =   "Form1.frx":8B993
      Top             =   960
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   19
      Left            =   6360
      Picture         =   "Form1.frx":9089D
      Top             =   720
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   18
      Left            =   6480
      Picture         =   "Form1.frx":957A7
      Top             =   960
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   17
      Left            =   6360
      Picture         =   "Form1.frx":9A6B1
      Top             =   720
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   16
      Left            =   6480
      Picture         =   "Form1.frx":9F5BB
      Top             =   960
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   15
      Left            =   6360
      Picture         =   "Form1.frx":A44C5
      Top             =   720
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   14
      Left            =   6480
      Picture         =   "Form1.frx":A93CF
      Top             =   960
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   13
      Left            =   6360
      Picture         =   "Form1.frx":AE2D9
      Top             =   720
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   12
      Left            =   6480
      Picture         =   "Form1.frx":B31E3
      Top             =   960
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   1200
      Index           =   11
      Left            =   6360
      Picture         =   "Form1.frx":B80ED
      Top             =   720
      Width           =   1200
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   10
      Left            =   6480
      Picture         =   "Form1.frx":BCFF7
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   9
      Left            =   6360
      Picture         =   "Form1.frx":BE091
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   8
      Left            =   6480
      Picture         =   "Form1.frx":BF12B
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   7
      Left            =   6360
      Picture         =   "Form1.frx":C01C5
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   6
      Left            =   6480
      Picture         =   "Form1.frx":C125F
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   5
      Left            =   6360
      Picture         =   "Form1.frx":C22F9
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   4
      Left            =   6480
      Picture         =   "Form1.frx":C3393
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   3
      Left            =   6360
      Picture         =   "Form1.frx":C442D
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   2
      Left            =   6480
      Picture         =   "Form1.frx":C54C7
      Top             =   960
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   540
      Index           =   1
      Left            =   6360
      Picture         =   "Form1.frx":C6561
      Top             =   720
      Width           =   540
   End
   Begin VB.Image 游戏图库 
      Height          =   375
      Index           =   0
      Left            =   6480
      Top             =   960
      Width           =   735
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   63
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   62
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   61
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   60
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   59
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   58
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   57
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   56
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   55
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   54
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   53
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   52
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   51
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   50
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   49
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   48
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   47
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   46
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   45
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   44
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   43
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   42
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   41
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   40
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   39
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   38
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   37
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   36
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   35
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   34
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   33
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   32
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   31
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   30
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   29
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   28
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   27
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   26
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   25
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   24
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   23
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   22
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   21
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   20
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   19
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   18
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   17
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   16
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   15
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   14
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   13
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   12
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   11
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   10
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   9
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   8
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   7
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   6
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   5
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   4
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   3
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   2
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   1
      Left            =   6600
      Top             =   720
      Width           =   615
   End
   Begin VB.Image 原始菜单图库 
      Height          =   495
      Index           =   0
      Left            =   6720
      Top             =   960
      Width           =   615
   End
   Begin VB.Image 流星 
      Height          =   375
      Index           =   7
      Left            =   6360
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 流星 
      Height          =   375
      Index           =   6
      Left            =   6480
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image 流星 
      Height          =   375
      Index           =   5
      Left            =   6360
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 流星 
      Height          =   375
      Index           =   4
      Left            =   6480
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image 流星 
      Height          =   375
      Index           =   3
      Left            =   6360
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 流星 
      Height          =   375
      Index           =   2
      Left            =   6480
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image 流星 
      Height          =   375
      Index           =   1
      Left            =   6360
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 流星 
      Height          =   375
      Index           =   0
      Left            =   6480
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image 桌面树 
      Height          =   615
      Left            =   6000
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   15
      Left            =   6360
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   14
      Left            =   6480
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   13
      Left            =   6360
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   12
      Left            =   6480
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   11
      Left            =   6360
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   10
      Left            =   6480
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   9
      Left            =   6360
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   8
      Left            =   6480
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   7
      Left            =   6360
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   6
      Left            =   6480
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   5
      Left            =   6360
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   4
      Left            =   6480
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   3
      Left            =   6360
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   2
      Left            =   6480
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   1
      Left            =   6360
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面粒子 
      Height          =   495
      Index           =   0
      Left            =   6480
      Top             =   960
      Width           =   495
   End
   Begin VB.Image 桌面星 
      Height          =   495
      Index           =   7
      Left            =   6600
      Top             =   720
      Width           =   495
   End
   Begin VB.Image 桌面星 
      Height          =   495
      Index           =   6
      Left            =   6720
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面星 
      Height          =   495
      Index           =   5
      Left            =   6600
      Top             =   720
      Width           =   495
   End
   Begin VB.Image 桌面星 
      Height          =   495
      Index           =   4
      Left            =   6720
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面星 
      Height          =   495
      Index           =   3
      Left            =   6600
      Top             =   720
      Width           =   495
   End
   Begin VB.Image 桌面星 
      Height          =   495
      Index           =   2
      Left            =   6720
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 桌面星 
      Height          =   495
      Index           =   1
      Left            =   6600
      Top             =   720
      Width           =   495
   End
   Begin VB.Image 桌面星 
      Height          =   495
      Index           =   0
      Left            =   6720
      Top             =   840
      Width           =   495
   End
   Begin VB.Image 菜单图库 
      Height          =   735
      Index           =   63
      Left            =   5760
      Top             =   480
      Width           =   735
   End
   Begin VB.Image 菜单图库 
      Height          =   735
      Index           =   62
      Left            =   5880
      Top             =   600
      Width           =   735
   End
   Begin VB.Image 菜单图库 
      Height          =   735
      Index           =   61
      Left            =   5760
      Top             =   480
      Width           =   735
   End
   Begin VB.Image 菜单图库 
      Height          =   735
      Index           =   60
      Left            =   5880
      Top             =   600
      Width           =   735
   End
   Begin VB.Image 菜单图库 
      Height          =   735
      Index           =   59
      Left            =   5760
      Top             =   480
      Width           =   735
   End
   Begin VB.Image 菜单图库 
      Height          =   735
      Index           =   58
      Left            =   5880
      Top             =   600
      Width           =   735
   End
   Begin VB.Image 菜单图库 
      Height          =   30
      Index           =   57
      Left            =   5760
      Picture         =   "Form1.frx":C75FB
      Top             =   480
      Width           =   30
   End
   Begin VB.Image 菜单图库 
      Height          =   1800
      Index           =   56
      Left            =   5880
      Picture         =   "Form1.frx":C765D
      Top             =   600
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   1800
      Index           =   55
      Left            =   5760
      Picture         =   "Form1.frx":D26E7
      Top             =   480
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   1800
      Index           =   54
      Left            =   5880
      Picture         =   "Form1.frx":DD771
      Top             =   600
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   30
      Index           =   53
      Left            =   5760
      Picture         =   "Form1.frx":E87FB
      Top             =   480
      Width           =   30
   End
   Begin VB.Image 菜单图库 
      Height          =   30
      Index           =   52
      Left            =   5880
      Picture         =   "Form1.frx":E885D
      Top             =   600
      Width           =   30
   End
   Begin VB.Image 菜单图库 
      Height          =   30
      Index           =   51
      Left            =   5760
      Picture         =   "Form1.frx":E88BF
      Top             =   480
      Width           =   30
   End
   Begin VB.Image 菜单图库 
      Height          =   120
      Index           =   50
      Left            =   5880
      Picture         =   "Form1.frx":E8921
      Top             =   600
      Width           =   120
   End
   Begin VB.Image 菜单图库 
      Height          =   120
      Index           =   49
      Left            =   5760
      Picture         =   "Form1.frx":E8A4B
      Top             =   480
      Width           =   120
   End
   Begin VB.Image 菜单图库 
      Height          =   120
      Index           =   48
      Left            =   5880
      Picture         =   "Form1.frx":E8B75
      Top             =   600
      Width           =   120
   End
   Begin VB.Image 菜单图库 
      Height          =   120
      Index           =   47
      Left            =   5760
      Picture         =   "Form1.frx":E8C9F
      Top             =   480
      Width           =   120
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   46
      Left            =   5880
      Picture         =   "Form1.frx":E8DC9
      Top             =   600
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   45
      Left            =   5760
      Picture         =   "Form1.frx":E9533
      Top             =   480
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   44
      Left            =   5880
      Picture         =   "Form1.frx":E9C9D
      Top             =   600
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   43
      Left            =   5760
      Picture         =   "Form1.frx":EA407
      Top             =   480
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   42
      Left            =   5880
      Picture         =   "Form1.frx":EAB71
      Top             =   600
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   41
      Left            =   5760
      Picture         =   "Form1.frx":EB2DB
      Top             =   480
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   40
      Left            =   5880
      Picture         =   "Form1.frx":EBA45
      Top             =   600
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   900
      Index           =   39
      Left            =   5760
      Picture         =   "Form1.frx":EC1AF
      Top             =   480
      Width           =   2700
   End
   Begin VB.Image 菜单图库 
      Height          =   900
      Index           =   38
      Left            =   5880
      Picture         =   "Form1.frx":F4629
      Top             =   600
      Width           =   2700
   End
   Begin VB.Image 菜单图库 
      Height          =   900
      Index           =   37
      Left            =   5760
      Picture         =   "Form1.frx":FCAA3
      Top             =   480
      Width           =   2700
   End
   Begin VB.Image 菜单图库 
      Height          =   900
      Index           =   36
      Left            =   5880
      Picture         =   "Form1.frx":104F1D
      Top             =   600
      Width           =   2700
   End
   Begin VB.Image 菜单图库 
      Height          =   900
      Index           =   35
      Left            =   5760
      Picture         =   "Form1.frx":10D397
      Top             =   480
      Width           =   2700
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   34
      Left            =   5880
      Picture         =   "Form1.frx":115811
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   33
      Left            =   5760
      Picture         =   "Form1.frx":11A71B
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   32
      Left            =   5880
      Picture         =   "Form1.frx":11F625
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   31
      Left            =   5760
      Picture         =   "Form1.frx":12452F
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   30
      Left            =   5880
      Picture         =   "Form1.frx":129439
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   29
      Left            =   5760
      Picture         =   "Form1.frx":12E343
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   28
      Left            =   5880
      Picture         =   "Form1.frx":13324D
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   27
      Left            =   5760
      Picture         =   "Form1.frx":138157
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   26
      Left            =   5880
      Picture         =   "Form1.frx":13D061
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   25
      Left            =   5760
      Picture         =   "Form1.frx":141F6B
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   24
      Left            =   5880
      Picture         =   "Form1.frx":146E75
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   23
      Left            =   5760
      Picture         =   "Form1.frx":14BD7F
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   22
      Left            =   5880
      Picture         =   "Form1.frx":150C89
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   21
      Left            =   5760
      Picture         =   "Form1.frx":155B93
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   20
      Left            =   5880
      Picture         =   "Form1.frx":15AA9D
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   19
      Left            =   5760
      Picture         =   "Form1.frx":15F9A7
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   18
      Left            =   5880
      Picture         =   "Form1.frx":1648B1
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   17
      Left            =   5760
      Picture         =   "Form1.frx":1697BB
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   16
      Left            =   5880
      Picture         =   "Form1.frx":16E6C5
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   15
      Left            =   5760
      Picture         =   "Form1.frx":1735CF
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   14
      Left            =   5880
      Picture         =   "Form1.frx":1784D9
      Top             =   600
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   1200
      Index           =   13
      Left            =   5760
      Picture         =   "Form1.frx":17D3E3
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   12
      Left            =   5880
      Picture         =   "Form1.frx":1822ED
      Top             =   600
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   11
      Left            =   5760
      Picture         =   "Form1.frx":184677
      Top             =   480
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   10
      Left            =   5880
      Picture         =   "Form1.frx":186A01
      Top             =   600
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   9
      Left            =   5760
      Picture         =   "Form1.frx":188D8B
      Top             =   480
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   8
      Left            =   5880
      Picture         =   "Form1.frx":18B115
      Top             =   600
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   7
      Left            =   5760
      Picture         =   "Form1.frx":18D49F
      Top             =   480
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   6
      Left            =   5880
      Picture         =   "Form1.frx":18F829
      Top             =   600
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   5
      Left            =   5760
      Picture         =   "Form1.frx":191BB3
      Top             =   480
      Width           =   1800
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   4
      Left            =   5880
      Picture         =   "Form1.frx":193F3D
      Top             =   600
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   3
      Left            =   5760
      Picture         =   "Form1.frx":1946A7
      Top             =   480
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   2
      Left            =   5880
      Picture         =   "Form1.frx":194E11
      Top             =   600
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   360
      Index           =   1
      Left            =   5760
      Picture         =   "Form1.frx":19557B
      Top             =   480
      Width           =   360
   End
   Begin VB.Image 菜单图库 
      Height          =   735
      Index           =   0
      Left            =   5880
      Top             =   600
      Width           =   735
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   23
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   22
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   21
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   20
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   19
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   18
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   17
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   960
      Index           =   16
      Left            =   6000
      Picture         =   "Form1.frx":195CE5
      Top             =   720
      Width           =   960
   End
   Begin VB.Label 起始文字 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   6000
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label 起始文字 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   6000
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label 起始文字 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6000
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label 起始文字 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   6720
      TabIndex        =   0
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   15
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   14
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   13
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   12
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   11
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   360
      Index           =   10
      Left            =   6000
      Picture         =   "Form1.frx":198F2F
      Top             =   720
      Width           =   1800
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   9
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   8
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   7
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   6
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   5
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   4
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   3
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   2
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   735
      Index           =   1
      Left            =   6000
      Top             =   720
      Width           =   975
   End
   Begin VB.Image 起始按钮 
      Height          =   240
      Index           =   0
      Left            =   6840
      Picture         =   "Form1.frx":19B2B9
      Top             =   3600
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   255
      Left            =   7200
      Picture         =   "Form1.frx":19B643
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   254
      Left            =   7080
      Picture         =   "Form1.frx":19BF0D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   253
      Left            =   7200
      Picture         =   "Form1.frx":19C7D7
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   252
      Left            =   7080
      Picture         =   "Form1.frx":19D0A1
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   251
      Left            =   7200
      Picture         =   "Form1.frx":19D96B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   250
      Left            =   7080
      Picture         =   "Form1.frx":19E235
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   249
      Left            =   7200
      Picture         =   "Form1.frx":19EAFF
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   248
      Left            =   7080
      Picture         =   "Form1.frx":19F3C9
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   247
      Left            =   7200
      Picture         =   "Form1.frx":19FC93
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   246
      Left            =   7080
      Picture         =   "Form1.frx":1A055D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   245
      Left            =   7200
      Picture         =   "Form1.frx":1A0E27
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   244
      Left            =   7080
      Picture         =   "Form1.frx":1A16F1
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   243
      Left            =   7200
      Picture         =   "Form1.frx":1A1FBB
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   242
      Left            =   7080
      Picture         =   "Form1.frx":1A2885
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   241
      Left            =   7200
      Picture         =   "Form1.frx":1A314F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   240
      Left            =   7080
      Picture         =   "Form1.frx":1A3A19
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   239
      Left            =   6120
      Picture         =   "Form1.frx":1A42E3
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   238
      Left            =   6000
      Picture         =   "Form1.frx":1A4BAD
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   237
      Left            =   6120
      Picture         =   "Form1.frx":1A5477
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   236
      Left            =   6000
      Picture         =   "Form1.frx":1A5D41
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   235
      Left            =   6120
      Picture         =   "Form1.frx":1A660B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   234
      Left            =   6000
      Picture         =   "Form1.frx":1A6ED5
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   233
      Left            =   6120
      Picture         =   "Form1.frx":1A779F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   232
      Left            =   6000
      Picture         =   "Form1.frx":1A8069
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   231
      Left            =   6120
      Picture         =   "Form1.frx":1A8933
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   230
      Left            =   6000
      Picture         =   "Form1.frx":1A91FD
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   229
      Left            =   6120
      Picture         =   "Form1.frx":1A9AC7
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   228
      Left            =   6000
      Picture         =   "Form1.frx":1AA391
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   227
      Left            =   6120
      Picture         =   "Form1.frx":1AAC5B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   226
      Left            =   6000
      Picture         =   "Form1.frx":1AB525
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   225
      Left            =   6120
      Picture         =   "Form1.frx":1ABDEF
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   224
      Left            =   6000
      Picture         =   "Form1.frx":1AC6B9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   223
      Left            =   7200
      Picture         =   "Form1.frx":1ACF83
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   222
      Left            =   7080
      Picture         =   "Form1.frx":1AD84D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   221
      Left            =   7200
      Picture         =   "Form1.frx":1AE117
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   220
      Left            =   7080
      Picture         =   "Form1.frx":1AE9E1
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   219
      Left            =   7200
      Picture         =   "Form1.frx":1AF2AB
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   218
      Left            =   7080
      Picture         =   "Form1.frx":1AFB75
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   217
      Left            =   7200
      Picture         =   "Form1.frx":1B043F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   216
      Left            =   7080
      Picture         =   "Form1.frx":1B0D09
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   215
      Left            =   7200
      Picture         =   "Form1.frx":1B15D3
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   214
      Left            =   7080
      Picture         =   "Form1.frx":1B1E9D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   213
      Left            =   7200
      Picture         =   "Form1.frx":1B2767
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   212
      Left            =   7080
      Picture         =   "Form1.frx":1B3031
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   211
      Left            =   7200
      Picture         =   "Form1.frx":1B38FB
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   210
      Left            =   7080
      Picture         =   "Form1.frx":1B41C5
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   209
      Left            =   7200
      Picture         =   "Form1.frx":1B4A8F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   208
      Left            =   7080
      Picture         =   "Form1.frx":1B5359
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   207
      Left            =   6120
      Picture         =   "Form1.frx":1B5C23
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   206
      Left            =   6000
      Picture         =   "Form1.frx":1B64ED
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   205
      Left            =   6120
      Picture         =   "Form1.frx":1B6DB7
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   204
      Left            =   6000
      Picture         =   "Form1.frx":1B7681
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   203
      Left            =   6120
      Picture         =   "Form1.frx":1B7F4B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   202
      Left            =   6000
      Picture         =   "Form1.frx":1B8815
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   201
      Left            =   6120
      Picture         =   "Form1.frx":1B90DF
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   200
      Left            =   6000
      Picture         =   "Form1.frx":1B99A9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   199
      Left            =   6120
      Picture         =   "Form1.frx":1BA273
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   198
      Left            =   6000
      Picture         =   "Form1.frx":1BAB3D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   197
      Left            =   6120
      Picture         =   "Form1.frx":1BB407
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   196
      Left            =   6000
      Picture         =   "Form1.frx":1BBCD1
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   195
      Left            =   6120
      Picture         =   "Form1.frx":1BC59B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   194
      Left            =   6000
      Picture         =   "Form1.frx":1BCE65
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   193
      Left            =   6120
      Picture         =   "Form1.frx":1BD72F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   192
      Left            =   6000
      Picture         =   "Form1.frx":1BDFF9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   191
      Left            =   7200
      Picture         =   "Form1.frx":1BE8C3
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   190
      Left            =   7080
      Picture         =   "Form1.frx":1BF18D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   189
      Left            =   7200
      Picture         =   "Form1.frx":1BFA57
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   188
      Left            =   7080
      Picture         =   "Form1.frx":1C0321
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   187
      Left            =   7200
      Picture         =   "Form1.frx":1C0BEB
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   186
      Left            =   7080
      Picture         =   "Form1.frx":1C14B5
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   185
      Left            =   7200
      Picture         =   "Form1.frx":1C1D7F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   184
      Left            =   7080
      Picture         =   "Form1.frx":1C2649
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   183
      Left            =   7200
      Picture         =   "Form1.frx":1C2F13
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   182
      Left            =   7080
      Picture         =   "Form1.frx":1C37DD
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   181
      Left            =   7200
      Picture         =   "Form1.frx":1C40A7
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   180
      Left            =   7080
      Picture         =   "Form1.frx":1C4971
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   179
      Left            =   7200
      Picture         =   "Form1.frx":1C523B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   178
      Left            =   7080
      Picture         =   "Form1.frx":1C5B05
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   177
      Left            =   7200
      Picture         =   "Form1.frx":1C63CF
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   176
      Left            =   7080
      Picture         =   "Form1.frx":1C6C99
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   175
      Left            =   6120
      Picture         =   "Form1.frx":1C7563
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   174
      Left            =   6000
      Picture         =   "Form1.frx":1C7E2D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   173
      Left            =   6120
      Picture         =   "Form1.frx":1C86F7
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   172
      Left            =   6000
      Picture         =   "Form1.frx":1C8FC1
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   171
      Left            =   6120
      Picture         =   "Form1.frx":1C988B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   170
      Left            =   6000
      Picture         =   "Form1.frx":1CA155
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   169
      Left            =   6120
      Picture         =   "Form1.frx":1CAA1F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   168
      Left            =   6000
      Picture         =   "Form1.frx":1CB2E9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   167
      Left            =   6120
      Picture         =   "Form1.frx":1CBBB3
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   166
      Left            =   6000
      Picture         =   "Form1.frx":1CC47D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   165
      Left            =   6120
      Picture         =   "Form1.frx":1CCD47
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   164
      Left            =   6000
      Picture         =   "Form1.frx":1CD611
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   163
      Left            =   6120
      Picture         =   "Form1.frx":1CDEDB
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   162
      Left            =   6000
      Picture         =   "Form1.frx":1CE7A5
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   161
      Left            =   6120
      Picture         =   "Form1.frx":1CF06F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   160
      Left            =   6000
      Picture         =   "Form1.frx":1CF939
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   159
      Left            =   7200
      Picture         =   "Form1.frx":1D0203
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   158
      Left            =   7080
      Picture         =   "Form1.frx":1D0ACD
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   157
      Left            =   7200
      Picture         =   "Form1.frx":1D1397
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   156
      Left            =   7080
      Picture         =   "Form1.frx":1D1C61
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   155
      Left            =   7200
      Picture         =   "Form1.frx":1D252B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   154
      Left            =   7080
      Picture         =   "Form1.frx":1D2DF5
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   153
      Left            =   7200
      Picture         =   "Form1.frx":1D36BF
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   152
      Left            =   7080
      Picture         =   "Form1.frx":1D3F89
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   151
      Left            =   7200
      Picture         =   "Form1.frx":1D4853
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   150
      Left            =   7080
      Picture         =   "Form1.frx":1D511D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   149
      Left            =   7200
      Picture         =   "Form1.frx":1D59E7
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   148
      Left            =   7080
      Picture         =   "Form1.frx":1D62B1
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   147
      Left            =   7200
      Picture         =   "Form1.frx":1D6B7B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   146
      Left            =   7080
      Picture         =   "Form1.frx":1D7445
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   145
      Left            =   7200
      Picture         =   "Form1.frx":1D7D0F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   144
      Left            =   7080
      Picture         =   "Form1.frx":1D85D9
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   143
      Left            =   6120
      Picture         =   "Form1.frx":1D8EA3
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   142
      Left            =   6000
      Picture         =   "Form1.frx":1D976D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   141
      Left            =   6120
      Picture         =   "Form1.frx":1DA037
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   140
      Left            =   6000
      Picture         =   "Form1.frx":1DA901
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   139
      Left            =   6120
      Picture         =   "Form1.frx":1DB1CB
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   138
      Left            =   6000
      Picture         =   "Form1.frx":1DBA95
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   137
      Left            =   6120
      Picture         =   "Form1.frx":1DC35F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   136
      Left            =   6000
      Picture         =   "Form1.frx":1DCC29
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   135
      Left            =   6120
      Picture         =   "Form1.frx":1DD4F3
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   134
      Left            =   6000
      Picture         =   "Form1.frx":1DDDBD
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   133
      Left            =   6120
      Picture         =   "Form1.frx":1DE687
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   132
      Left            =   6000
      Picture         =   "Form1.frx":1DEF51
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   131
      Left            =   6120
      Picture         =   "Form1.frx":1DF81B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   130
      Left            =   6000
      Picture         =   "Form1.frx":1E00E5
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   129
      Left            =   6120
      Picture         =   "Form1.frx":1E09AF
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   128
      Left            =   6000
      Picture         =   "Form1.frx":1E1279
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   127
      Left            =   7200
      Picture         =   "Form1.frx":1E1B43
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   126
      Left            =   7080
      Picture         =   "Form1.frx":1E240D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   125
      Left            =   7200
      Picture         =   "Form1.frx":1E2CD7
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   124
      Left            =   7080
      Picture         =   "Form1.frx":1E35A1
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   123
      Left            =   7200
      Picture         =   "Form1.frx":1E3E6B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   122
      Left            =   7080
      Picture         =   "Form1.frx":1E4735
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   121
      Left            =   7200
      Picture         =   "Form1.frx":1E4FFF
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   120
      Left            =   7080
      Picture         =   "Form1.frx":1E58C9
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   119
      Left            =   7200
      Picture         =   "Form1.frx":1E6193
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   118
      Left            =   7080
      Picture         =   "Form1.frx":1E6A5D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   117
      Left            =   7200
      Picture         =   "Form1.frx":1E7327
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   116
      Left            =   7080
      Picture         =   "Form1.frx":1E7BF1
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   115
      Left            =   7200
      Picture         =   "Form1.frx":1E84BB
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   114
      Left            =   7080
      Picture         =   "Form1.frx":1E8D85
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   113
      Left            =   7200
      Picture         =   "Form1.frx":1E964F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   112
      Left            =   7080
      Picture         =   "Form1.frx":1E9F19
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   111
      Left            =   6120
      Picture         =   "Form1.frx":1EA7E3
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   110
      Left            =   6000
      Picture         =   "Form1.frx":1EB0AD
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   109
      Left            =   6120
      Picture         =   "Form1.frx":1EB977
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   108
      Left            =   6000
      Picture         =   "Form1.frx":1EC241
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   107
      Left            =   6120
      Picture         =   "Form1.frx":1ECB0B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   106
      Left            =   6000
      Picture         =   "Form1.frx":1ED3D5
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   105
      Left            =   6120
      Picture         =   "Form1.frx":1EDC9F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   104
      Left            =   6000
      Picture         =   "Form1.frx":1EE569
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   103
      Left            =   6120
      Picture         =   "Form1.frx":1EEE33
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   102
      Left            =   6000
      Picture         =   "Form1.frx":1EF6FD
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   101
      Left            =   6120
      Picture         =   "Form1.frx":1EFFC7
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   100
      Left            =   6000
      Picture         =   "Form1.frx":1F0891
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   99
      Left            =   6120
      Picture         =   "Form1.frx":1F115B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   98
      Left            =   6000
      Picture         =   "Form1.frx":1F1A25
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   97
      Left            =   6120
      Picture         =   "Form1.frx":1F22EF
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   96
      Left            =   6000
      Picture         =   "Form1.frx":1F2BB9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   95
      Left            =   7200
      Picture         =   "Form1.frx":1F3483
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   94
      Left            =   7080
      Picture         =   "Form1.frx":1F3D4D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   93
      Left            =   7200
      Picture         =   "Form1.frx":1F4617
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   92
      Left            =   7080
      Picture         =   "Form1.frx":1F4EE1
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   91
      Left            =   7200
      Picture         =   "Form1.frx":1F57AB
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   90
      Left            =   7080
      Picture         =   "Form1.frx":1F6075
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   89
      Left            =   7200
      Picture         =   "Form1.frx":1F693F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   88
      Left            =   7080
      Picture         =   "Form1.frx":1F7209
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   87
      Left            =   7200
      Picture         =   "Form1.frx":1F7AD3
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   86
      Left            =   7080
      Picture         =   "Form1.frx":1F839D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   85
      Left            =   7200
      Picture         =   "Form1.frx":1F8C67
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   84
      Left            =   7080
      Picture         =   "Form1.frx":1F9531
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   83
      Left            =   7200
      Picture         =   "Form1.frx":1F9DFB
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   82
      Left            =   7080
      Picture         =   "Form1.frx":1FA6C5
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   81
      Left            =   7200
      Picture         =   "Form1.frx":1FAF8F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   80
      Left            =   7080
      Picture         =   "Form1.frx":1FB859
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   79
      Left            =   6120
      Picture         =   "Form1.frx":1FC123
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   78
      Left            =   6000
      Picture         =   "Form1.frx":1FC9ED
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   77
      Left            =   6120
      Picture         =   "Form1.frx":1FD2B7
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   76
      Left            =   6000
      Picture         =   "Form1.frx":1FDB81
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   75
      Left            =   6120
      Picture         =   "Form1.frx":1FE44B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   74
      Left            =   6000
      Picture         =   "Form1.frx":1FED15
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   73
      Left            =   6120
      Picture         =   "Form1.frx":1FF5DF
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   72
      Left            =   6000
      Picture         =   "Form1.frx":1FFEA9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   71
      Left            =   6120
      Picture         =   "Form1.frx":200773
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   70
      Left            =   6000
      Picture         =   "Form1.frx":20103D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   69
      Left            =   6120
      Picture         =   "Form1.frx":201907
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   68
      Left            =   6000
      Picture         =   "Form1.frx":2021D1
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   67
      Left            =   6120
      Picture         =   "Form1.frx":202A9B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   66
      Left            =   6000
      Picture         =   "Form1.frx":203365
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   65
      Left            =   6120
      Picture         =   "Form1.frx":203C2F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   64
      Left            =   6000
      Picture         =   "Form1.frx":2044F9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   63
      Left            =   7200
      Picture         =   "Form1.frx":204DC3
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   62
      Left            =   7080
      Picture         =   "Form1.frx":20568D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   61
      Left            =   7200
      Picture         =   "Form1.frx":205F57
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   60
      Left            =   7080
      Picture         =   "Form1.frx":206821
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   59
      Left            =   7200
      Picture         =   "Form1.frx":2070EB
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   58
      Left            =   7080
      Picture         =   "Form1.frx":2079B5
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   57
      Left            =   7200
      Picture         =   "Form1.frx":20827F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   56
      Left            =   7080
      Picture         =   "Form1.frx":208B49
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   55
      Left            =   7200
      Picture         =   "Form1.frx":209413
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   54
      Left            =   7080
      Picture         =   "Form1.frx":209CDD
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   53
      Left            =   7200
      Picture         =   "Form1.frx":20A5A7
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   52
      Left            =   7080
      Picture         =   "Form1.frx":20AE71
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   51
      Left            =   7200
      Picture         =   "Form1.frx":20B73B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   50
      Left            =   7080
      Picture         =   "Form1.frx":20C005
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   49
      Left            =   7200
      Picture         =   "Form1.frx":20C8CF
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   48
      Left            =   7080
      Picture         =   "Form1.frx":20D199
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   47
      Left            =   6120
      Picture         =   "Form1.frx":20DA63
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   46
      Left            =   6000
      Picture         =   "Form1.frx":20E32D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   45
      Left            =   6120
      Picture         =   "Form1.frx":20EBF7
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   44
      Left            =   6000
      Picture         =   "Form1.frx":20F4C1
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   43
      Left            =   6120
      Picture         =   "Form1.frx":20FD8B
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   42
      Left            =   6000
      Picture         =   "Form1.frx":210655
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   41
      Left            =   6120
      Picture         =   "Form1.frx":210F1F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   40
      Left            =   6000
      Picture         =   "Form1.frx":2117E9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   39
      Left            =   6120
      Picture         =   "Form1.frx":2120B3
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   38
      Left            =   6000
      Picture         =   "Form1.frx":21297D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   37
      Left            =   6120
      Picture         =   "Form1.frx":213247
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   36
      Left            =   6000
      Picture         =   "Form1.frx":213B11
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   35
      Left            =   6120
      Picture         =   "Form1.frx":2143DB
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   34
      Left            =   6000
      Picture         =   "Form1.frx":214CA5
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   33
      Left            =   6120
      Picture         =   "Form1.frx":21556F
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   32
      Left            =   6000
      Picture         =   "Form1.frx":215E39
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   31
      Left            =   7200
      Picture         =   "Form1.frx":216703
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   30
      Left            =   7080
      Picture         =   "Form1.frx":216FCD
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   29
      Left            =   7200
      Picture         =   "Form1.frx":217897
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   28
      Left            =   7080
      Picture         =   "Form1.frx":218161
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   27
      Left            =   7200
      Picture         =   "Form1.frx":218A2B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   26
      Left            =   7080
      Picture         =   "Form1.frx":2192F5
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   25
      Left            =   7200
      Picture         =   "Form1.frx":219BBF
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   24
      Left            =   7080
      Picture         =   "Form1.frx":21A489
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   23
      Left            =   7200
      Picture         =   "Form1.frx":21AD53
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   22
      Left            =   7080
      Picture         =   "Form1.frx":21B61D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   21
      Left            =   7200
      Picture         =   "Form1.frx":21BEE7
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   20
      Left            =   7080
      Picture         =   "Form1.frx":21C7B1
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   19
      Left            =   7200
      Picture         =   "Form1.frx":21D07B
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   18
      Left            =   7080
      Picture         =   "Form1.frx":21D945
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   17
      Left            =   7200
      Picture         =   "Form1.frx":21E20F
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   16
      Left            =   7080
      Picture         =   "Form1.frx":21EAD9
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   15
      Left            =   6000
      Picture         =   "Form1.frx":21F3A3
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   14
      Left            =   6120
      Picture         =   "Form1.frx":21FC6D
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   13
      Left            =   6000
      Picture         =   "Form1.frx":220537
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   12
      Left            =   6120
      Picture         =   "Form1.frx":220E01
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   11
      Left            =   6000
      Picture         =   "Form1.frx":2216CB
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   10
      Left            =   6120
      Picture         =   "Form1.frx":221F95
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   9
      Left            =   6000
      Picture         =   "Form1.frx":22285F
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   8
      Left            =   6120
      Picture         =   "Form1.frx":223129
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   7
      Left            =   6000
      Picture         =   "Form1.frx":2239F3
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   6
      Left            =   6120
      Picture         =   "Form1.frx":2242BD
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   5
      Left            =   6000
      Picture         =   "Form1.frx":224B87
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   4
      Left            =   6120
      Picture         =   "Form1.frx":225451
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   3
      Left            =   6000
      Picture         =   "Form1.frx":225D1B
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   2
      Left            =   6120
      Picture         =   "Form1.frx":2265E5
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   1
      Left            =   6000
      Picture         =   "Form1.frx":226EAF
      Stretch         =   -1  'True
      Top             =   720
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   0
      Left            =   6120
      Picture         =   "Form1.frx":227779
      Stretch         =   -1  'True
      Top             =   960
      Width           =   600
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X(201) As Long 'top
Dim Y(201) As Long 'left
Dim ax(12, 4) As Long, ay(12, 4) As Long ''一个四合方块组中各个方块相对于中心方块的位置
Dim b(5) As Long '砖块四组合分别在第几个方块
Dim 出新 As Boolean, 可下降 As Boolean, 可左移 As Boolean, 可右移 As Boolean, timing As Boolean
Dim li(11) As Long
Dim bb As Long
Dim formmousex As Long, formmousey As Long '用来记录最后form mouse位置来区别数组控件的的不同小控件
Dim 桌面树时间 As Long
Dim 桌面树果实 As Boolean
Dim 按键(5) As Long
Dim 起始文字数组(4) As Long
Dim 开始按钮动画位置 As Long
Dim 开始按钮进行中 As Boolean
Dim 人名(16) As String
Dim 能量变量 As Long
Dim 变脸时间 As Long
Dim 电脑选人(3) As Long
Dim 紧张时间 As Long
Dim 正在加速 As Boolean
Dim 快的速度 As Long, 慢的速度 As Long
Dim 能量小偷力量 As Long
Dim 电脑等级变量 As Long, 玩家等级变量 As Long
Dim 已经左右切换 As Boolean
Dim 左右切换间隔 As Long
Dim 永恒之日变量 As Long
Dim 材料小偷变量 As Long
Dim 树木皇后变量 As Long
Dim 键盘不让用 As Boolean
Dim 消失变量 As Long
Dim 成功 As Long
Dim 经验增量 As Long
Dim 金钱增量 As Long
Dim 该回菜单 As Boolean
Dim 宠物已经购买(5) As Boolean '1,2,3,4
Dim 正在游戏 As Boolean
Dim 奖励剩余时间(51) As Long
Dim 奖励指针 As Long
Dim 紧张延长 As Long
Sub 虎威()
If Int(Rnd * 10) + 1 <= 5 Then
Dim hu As Long
For hu = 1 To 200
Image1(hu).Visible = False
Next
End If
End Sub
 Sub 左键过程()
    Dim f As Long '【左移】开始
    可左移 = True
    For f = 1 To 4
     Image1(b(f)).Visible = False '排除原图形的干扰
    Next
    For f = 1 To 4
     If (Image1(b(f) - 1).Visible = True) Or (b(f) Mod 10 = 1) Then '寻找不可左移条件是否成立
     可左移 = False
     End If
    Next
    If 可左移 Then
      For f = 1 To 4
        b(f) = b(f) - 1
        Image1(b(f)).Visible = True '正在左移
      Next
    Else
     For f = 1 To 4
      Image1(b(f)).Visible = True '“排除原图形的干扰”的还原
     Next
    End If '【左移】结束
 End Sub
 Sub 右键过程()
     'Dim f As Long '【右移】开始
    可右移 = True
    For f = 1 To 4
     Image1(b(f)).Visible = False '排除原图形的干扰
    Next
    For f = 1 To 4
     If (Image1(b(f) + 1).Visible = True) Or (b(f) Mod 10 = 0) Then '寻找不可右移条件是否成立
     可右移 = False
     End If
    Next
    If 可右移 Then
      For f = 1 To 4
        b(f) = b(f) + 1
        Image1(b(f)).Visible = True '正在右移
      Next
    Else
     For f = 1 To 4
      Image1(b(f)).Visible = True '“排除原图形的干扰”的还原
     Next
    End If '【右移】结束
 End Sub

Private Sub Form_Click()
If 该回菜单 = True Then
该回菜单 = False
'=====================================回到主菜单====================================================
正在游戏 = False
Dim qs As Long
For qs = 0 To 7
桌面星(qs).Visible = True
流星(qs).Visible = True
Next
For qs = 0 To 15
桌面粒子(qs).Visible = True
Next
 For qs = 1 To 23
 起始按钮(qs).Visible = True
 Next
 For qs = 1 To 3
 起始文字(qs).Visible = True
 Next
 桌面树.Visible = True
'经验图标布局：
With 起始按钮(17)
.Visible = True
.Top = 200
.Left = 2200
.Picture = 菜单图库(41).Picture
End With
'金钱图标布局：
With 起始按钮(18)
.Visible = True
.Top = 200
.Left = 4100
.Picture = 菜单图库(42).Picture
End With
'菜单界面数字布局：
起始文字数组(1) = 起始文字数组(1)
起始文字数组(2) = 起始文字数组(2) + 经验增量
起始文字数组(3) = 起始文字数组(3) + 金钱增量
For qs = 1 To 3
With 起始文字(qs)
.Caption = 起始文字数组(qs)
.Top = 200
'‘’‘’‘’‘’‘’‘'小二字体
End With
Next
起始文字(1).Left = 1000
起始文字(2).Left = 2800
起始文字(3).Left = 4600
结果(0).Visible = False
结果(1).Visible = False
开始按钮动画时间系统.Enabled = False
For qs = 1 To 4
按键(qs) = 0
Next
起始按钮(15).Picture = 菜单图库(35).Picture



'===================================================================================================
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If timing = False And 键盘不让用 = False And b(1) <> 0 Then
Dim f As Long
'======================================================================
If KeyCode = vbKeyQ And 能量变量 >= 10 Then
能量变量 = 能量变量 - 10
能量数字.Caption = 能量变量
快的速度 = 快的速度 + 2
慢的速度 = 慢的速度 + 50
End If
If KeyCode = vbKeyW And 能量变量 >= 15 Then
能量变量 = 能量变量 - 15
能量数字.Caption = 能量变量
Dim lou As Long
For lou = 191 To 200
Image1(lou).Visible = True
Next
End If
If KeyCode = vbKeyE And 能量变量 >= 20 Then
能量变量 = 能量变量 - 20
能量数字.Caption = 能量变量
虎威
End If
If KeyCode = vbKeyS Then
For f = 1 To 200
Image1(f).Visible = False '==================================================必修改
Next
End If
If KeyCode = vbKeyUp Then
  For f = 1 To 4
    Image1(b(f)).Visible = False
  Next
  Dim 可变换 As Boolean
  Dim xi As Long, yi As Long, xii As Long, yii As Long
  可变换 = True
  xi = b(1) \ 10 + 1 'b(1)在第几行
  yi = b(1) Mod 10 'b(1)在第几列
  If yi = 0 Then yi = 10
  For f = 1 To 4 '检验b（1）周围三个方块是否有位置可以变换
    xii = ax(li(bb), f) + xi
    yii = ay(li(bb), f) + yi
    If (Image1(10 * (xii - 1) + yii).Visible = True) Or (xii > 20) Or (xii < 1) Or (yii > 10) Or (yii < 1) Then 可变换 = False
  Next
  If 可变换 Then
    For f = 1 To 3
      xii = ax(li(bb), f) + xi
      yii = ay(li(bb), f) + yi
      b(f + 1) = 10 * (xii - 1) + yii
    Next
    For f = 1 To 4
      Image1(b(f)).Visible = True
    Next
    bb = li(bb)
  Else
    For f = 1 To 4
    Image1(b(f)).Visible = True
    Next
  End If
End If
'=======================================================================
If KeyCode = vbKeyRight Then
If 已经左右切换 Then
左键过程
Else
右键过程
End If
End If
If KeyCode = vbKeyLeft Then
If 已经左右切换 Then
右键过程
Else
左键过程
End If
End If
If KeyCode = vbKeyDown Then
   '奖励
   能量变量 = 能量变量 + 1
   能量数字.Caption = 能量变量
          '奖励图标
  
  With 奖励(奖励指针)
  .Picture = 游戏图库(31).Picture
  .Left = Image1(b(1)).Left
  .Top = Image1(b(1)).Top
  End With
  奖励剩余时间(奖励指针) = 50
    奖励指针 = 奖励指针 + 1
    If 奖励指针 > 50 Then 奖励指针 = 1

    'Dim f As Long '【下降】开始
    可下降 = True
    For f = 1 To 4
     Image1(b(f)).Visible = False '排除原图形的干扰
    Next
    For f = 1 To 4
     If (Image1(b(f) + 10).Visible = True) Or (b(f) + 10 > 200) Then '寻找不可下降条件是否成立
     可下降 = False
     End If
    Next
    If 可下降 Then
      For f = 1 To 4
        b(f) = b(f) + 10
        Image1(b(f)).Visible = True '正在下降
      Next
    Else
     For f = 1 To 4
      Image1(b(f)).Visible = True '“排除原图形的干扰”的还原
     Next
     出新 = True
     timing = True
    End If '【下降】结束
        'Dim f As Long '【下降】开始
    可下降 = True
    For f = 1 To 4
     Image1(b(f)).Visible = False '排除原图形的干扰
    Next
    For f = 1 To 4
     If (Image1(b(f) + 10).Visible = True) Or (b(f) + 10 > 200) Then '寻找不可下降条件是否成立
     可下降 = False
     End If
    Next
    If 可下降 Then
      For f = 1 To 4
        b(f) = b(f) + 10
        Image1(b(f)).Visible = True '正在下降
      Next
    Else
     For f = 1 To 4
      Image1(b(f)).Visible = True '“排除原图形的干扰”的还原
     Next
     出新 = True
     timing = True
    End If '【下降】结束
        'Dim f As Long '【下降】开始
    可下降 = True
    For f = 1 To 4
     Image1(b(f)).Visible = False '排除原图形的干扰
    Next
    For f = 1 To 4
     If (Image1(b(f) + 10).Visible = True) Or (b(f) + 10 > 200) Then '寻找不可下降条件是否成立
     可下降 = False
     End If
    Next
    If 可下降 Then
      For f = 1 To 4
        b(f) = b(f) + 10
        Image1(b(f)).Visible = True '正在下降
      Next
    Else
     For f = 1 To 4
      Image1(b(f)).Visible = True '“排除原图形的干扰”的还原
     Next
     出新 = True
     timing = True
    End If '【下降】结束
End If
End If
End Sub

Private Sub Form_Load()
 With 提示不同能力
'form直接设定

End With
'----------------------------
紧张延长 = 0
'奖励图标设置
奖励指针 = 1
奖励时间系统.Interval = 15
奖励时间系统.Enabled = True
Dim jl As Long
For jl = 1 To 50
奖励(jl).Stretch = False '显示原图标大小
奖励剩余时间(jl) = 0
Next
'===================
正在游戏 = False
'===============
Dim cho As Long
'菜单文字背景透明
For cho = 1 To 3
起始文字(cho).BackStyle = 0
Next
'====

For cho = 1 To 4
宠物已经购买(cho) = False
Next
起始按钮(1).ToolTipText = "生成txt储存游戏（功能未开启）"
起始按钮(10).ToolTipText = "不吃东西可以活命"
起始按钮(11).ToolTipText = "猫吃5元猫食，购买宠物可用"
起始按钮(12).ToolTipText = "虎吃6元火腿，购买宠物可用"
起始按钮(13).ToolTipText = "猴吃7元香蕉，购买宠物可用"
起始按钮(14).ToolTipText = "它吃元神，算你8元，购买宠物可用"
起始按钮(19).ToolTipText = "敌人数量"
'升级界面：
With 升级界面
.Visible = False
.Height = 10600
.Width = 8000
.Top = 100
.Left = 100
End With
'商店界面
With 商店
.Visible = False
.Stretch = False
.Top = 200
.Left = 0
End With
买招财猫.Visible = False
买虎不发威.Visible = False
买学猴王.Visible = False
买精力十足.Visible = False
'人名加载：
人名(1) = "能量小偷"
人名(2) = "左右不分"
人名(3) = "流行之王"
人名(4) = "永恒之日"
人名(5) = "材料小偷"
人名(6) = "树木皇后"
人名(7) = "变异蜗牛"
人名(8) = "静心圣人"
人名(9) = "能量领袖"
人名(10) = "变通公主"
人名(11) = "不忙乌龟"
人名(12) = "招财小猫"
人名(13) = "虎不发威"
人名(14) = "学霸小猴"
人名(15) = "精力十足"
Dim k6 As Long
'菜单图库备用,但应该用不到：
For k6 = 0 To 63
原始菜单图库(k6).Picture = 菜单图库(k6).Picture
原始菜单图库(k6).Visible = False
Next
'为啥不管用：
'form1.Caption = " ff"因为后面还有一个
'时间速度、变量赋值：
游戏变脸时间系统.Interval = 1000
桌面星时间系统.Interval = 500
桌面粒子时间系统.Interval = 1
桌面树时间系统.Interval = 1000
流星时间系统.Interval = 200
桌面树时间 = 0
开始按钮动画时间系统.Interval = 150
开始按钮进行中 = False
紧张时间系统.Interval = 1000
该回菜单 = False
'下降时间不可用：
timer1.Enabled = False
'开始按钮动画时间不可用：
开始按钮动画时间系统.Enabled = False
'=========================================================================菜单界面打开==================================================================
'菜单按键记录初始化：
For k6 = 1 To 4
按键(k6) = 0
Next
'初始化全部图像不可见：
结果(0).Visible = False
结果(1).Visible = False
For k6 = 0 To 63
菜单图库(k6).Visible = False
Next
For k6 = 0 To 3
起始文字(k6).Visible = False
Next
For k6 = 0 To 23
起始按钮(k6).Visible = False
起始按钮(k6).Stretch = False '按真实图像大小显示
Next
For k6 = 1 To 200 '未验证
Image1(k6).Visible = False
Next
玩家一.Visible = False
玩家二.Visible = False
电脑一.Visible = False
电脑二.Visible = False
玩家等级.Visible = False
电脑等级.Visible = False
玩家名.Visible = False
电脑名.Visible = False
游戏时间.Visible = False
游戏时间数字.Visible = False
能量.Visible = False
能量数字.Visible = False
q键.Visible = False
w键.Visible = False
e键.Visible = False
For k6 = 0 To 63
游戏图库(k6).Visible = False
Next
'等级图标布局：
起始按钮(16).Visible = True
起始按钮(16).Top = 200
起始按钮(16).Left = 300
'起始按钮(16).Height = 360'提示大小作用（后面）
'起始按钮(16).Width = 360
起始按钮(16).Picture = 菜单图库(40).Picture
'经验图标布局：
With 起始按钮(17)
.Visible = True
.Top = 200
.Left = 2200
.Picture = 菜单图库(41).Picture
End With
'金钱图标布局：
With 起始按钮(18)
.Visible = True
.Top = 200
.Left = 4100
.Picture = 菜单图库(42).Picture
End With
'历史记录图标布局：
With 起始按钮(1)
.Visible = True
.Top = 200
.Left = 6300
.Picture = 菜单图库(1).Picture
End With
'商店图标布局：
With 起始按钮(2)
.Visible = True
.Top = 200
.Left = 7300
.Picture = 菜单图库(3).Picture
End With
'敌人图标布局：
With 起始按钮(19)
.Visible = True
.Top = 1500
.Left = 300
.Picture = 菜单图库(43).Picture
End With
'难度图标布局：
With 起始按钮(20)
.Visible = True
.Top = 2100
.Left = 300
.Picture = 菜单图库(44).Picture
End With
'人物图标布局：
With 起始按钮(21)
.Visible = True
.Top = 2700
.Left = 300
.Picture = 菜单图库(45).Picture
End With
'宠物图标布局：
With 起始按钮(22)
.Visible = True
.Top = 4340
.Left = 300
.Picture = 菜单图库(46).Picture
End With
'一个图标布局
With 起始按钮(3)
.Visible = True
.Top = 1500
.Left = 900
.Width = 1800 '提示大小作用（后面）
.Height = 360 '提示大小作用（后面）
.Picture = 菜单图库(5).Picture
End With
'两个图标布局
With 起始按钮(4)
.Visible = True
.Top = 1500
.Left = 3700
.Picture = 菜单图库(7).Picture
End With
'简单图标布局
With 起始按钮(5)
.Visible = True
.Top = 2100
.Left = 900
.Picture = 菜单图库(9).Picture
End With
'一般图标布局
With 起始按钮(6)
.Visible = True
.Top = 2100
.Left = 3700
.Picture = 菜单图库(11).Picture
End With
'静心圣人图标布局
With 起始按钮(7)
.Visible = True
.Top = 2700
.Left = 900
.Height = 1200
.Width = 1200
.Picture = 菜单图库(13).Picture
End With
'能量领袖图标布局
With 起始按钮(8)
.Visible = True
.Top = 2700
.Left = 3100
.Picture = 菜单图库(17).Picture
End With
'变通公主
With 起始按钮(9)
.Visible = True
.Top = 2700
.Left = 5300
.Picture = 菜单图库(21).Picture
End With
'乌龟图标布局
With 起始按钮(10)
.Visible = True
.Top = 4340
.Left = 900
.Picture = 菜单图库(25).Picture
End With
'招财猫图标布局
With 起始按钮(11)
.Visible = True
.Top = 4340
.Left = 3100
.Picture = 菜单图库(27).Picture
End With
'虎不发威图标布局：
With 起始按钮(12)
.Visible = True
.Top = 4340
.Left = 5300
.Picture = 菜单图库(29).Picture
End With
'学猴王图标布局：
With 起始按钮(13)
.Visible = True
.Top = 5740
.Left = 900
.Picture = 菜单图库(31).Picture
End With
'精力十足图标布局：
With 起始按钮(14)
.Visible = True
.Top = 5740
.Left = 3100
.Picture = 菜单图库(33).Picture
End With
'开始图标布局：
With 起始按钮(15)
.Visible = True
.Top = 9300
.Left = 2400
.Picture = 菜单图库(35).Picture
End With
'桌面星星布局：
桌面星时间系统.Enabled = True
For k6 = 1 To 6
桌面星(k6).Visible = True
桌面星(k6).Stretch = False
Next
With 桌面星(1)
.Top = 1700
.Left = 6400
.Picture = 菜单图库(47).Picture
End With
With 桌面星(2)
.Top = 10000
.Left = 7000
.Picture = 菜单图库(47).Picture
End With
With 桌面星(3)
.Top = 600
.Left = 800
.Picture = 菜单图库(47).Picture
End With
With 桌面星(4)
.Top = 3900
.Left = 7400
.Picture = 菜单图库(47).Picture
End With
With 桌面星(5)
.Top = 9000
.Left = 300
.Picture = 菜单图库(47).Picture
End With
With 桌面星(6)
.Top = 6700
.Left = 4400
.Picture = 菜单图库(47).Picture
End With
'桌面粒子图标布局：
For k6 = 1 To 15
桌面粒子(k6).Visible = True
桌面粒子(k6).Stretch = False
桌面粒子(k6).Top = 11300
桌面粒子(k6).Left = 8800
桌面粒子(k6).Picture = 菜单图库(51).Picture
Next
'桌面树图标布局
With 桌面树
.Visible = True
.Top = 7000
.Left = 6000
.Picture = 菜单图库(54).Picture
桌面树果实 = True
End With
'流星图标布局
For k6 = 1 To 7
With 流星(k6)
.Visible = True
.Picture = 菜单图库(57).Picture
.Top = -1000 - k6 * 1000
.Left = k6 * 1000
End With
Next
'菜单界面数字布局：
起始文字数组(1) = 1
起始文字数组(2) = 0
起始文字数组(3) = 0 '以前100
For k6 = 1 To 3
With 起始文字(k6)
.Visible = True
.Caption = 起始文字数组(k6)
.BackColor = &H0&     '黑
.ForeColor = &HFFFFFF  '文字白色
.Height = 360
.Width = 1000
.Top = 200
'‘’‘’‘’‘’‘’‘'小二字体
End With
Next
起始文字(1).Left = 1000
起始文字(2).Left = 2800
起始文字(3).Left = 4600

'=======================================================================================================================================================
'’‘’‘’‘’创建0--255个image1（i）并使图像大小随控件改变  创建一个timer
'‘’‘’‘’‘创建起始按钮（0~23）【image控件】
'’‘’‘’‘’创建起始文字（0~3）【label控件】
'‘’‘’‘’‘导入各个image的图片
'方块组之间的变化关系：
li(1) = 4
li(4) = 5
li(5) = 7
li(7) = 1
li(2) = 3
li(3) = 2
li(6) = 6
li(8) = 9
li(9) = 8
li(10) = 11
li(11) = 10
'一个四合方块组中各个方块相对于中心方块的位置：
ax(1, 1) = 1
ay(1, 1) = 0
ax(1, 2) = 1
ay(1, 2) = 1
ax(1, 3) = 2
ay(1, 3) = 0
ax(2, 1) = 1
ay(2, 1) = 0
ax(2, 2) = 2
ay(2, 2) = 0
ax(2, 3) = 3
ay(2, 3) = 0
ax(3, 1) = 0
ay(3, 1) = -1
ax(3, 2) = 0
ay(3, 2) = 1
ax(3, 3) = 0
ay(3, 3) = 2
ax(4, 1) = 1
ay(4, 1) = -1
ax(4, 2) = 1
ay(4, 2) = 0
ax(4, 3) = 1
ay(4, 3) = 1
ax(5, 1) = 1
ay(5, 1) = -1
ax(5, 2) = 1
ay(5, 2) = 0
ax(5, 3) = 2
ay(5, 3) = 0
ax(6, 1) = 0
ay(6, 1) = -1
ax(6, 2) = 1
ay(6, 2) = -1
ax(6, 3) = 1
ay(6, 3) = 0
ax(7, 1) = 0
ay(7, 1) = -1
ax(7, 2) = 1
ay(7, 2) = 0
ax(7, 3) = 0
ay(7, 3) = 1
ax(8, 1) = 1
ay(8, 1) = 0
ax(8, 2) = 1
ay(8, 2) = 1
ax(8, 3) = 2
ay(8, 3) = 1
ax(9, 1) = 1
ay(9, 1) = -1
ax(9, 2) = 1
ay(9, 2) = 0
ax(9, 3) = 0
ay(9, 3) = 1
ax(10, 1) = 1
ay(10, 1) = -1
ax(10, 2) = 1
ay(10, 2) = 0
ax(10, 3) = 2
ay(10, 3) = -1
ax(11, 1) = 0
ay(11, 1) = -1
ax(11, 2) = 1
ay(11, 2) = 0
ax(11, 3) = 1
ay(11, 3) = 1
form1.Height = 11400
form1.Width = 8400
Dim i As Long
For i = 1 To 200
Image1(i).Stretch = True
Image1(i).Height = 540
Image1(i).Width = 540
'image1（i=1..200）.pictrue=
X(i) = 540 * ((i - 1) \ 10)
Image1(i).Top = X(i)
Y(i) = 540 * ((i - 1) Mod 10)
Image1(i).Left = Y(i)
Next
For i = 0 To 255
Image1(i).Visible = False
Next
出新 = True
'form1.Caption = ""
timer1.Interval = 1000





'隐藏玩家无法选择的按钮
If 起始文字数组(1) >= 5 Then
起始按钮(8).Visible = True
Else
起始按钮(8).Visible = False
End If
If 起始文字数组(1) >= 10 Then
起始按钮(6).Visible = True
Else
起始按钮(6).Visible = False
End If
If 起始文字数组(1) >= 15 Then
起始按钮(9).Visible = True
Else
起始按钮(9).Visible = False
End If
If 宠物已经购买(1) Then
起始按钮(11).Visible = True
Else
起始按钮(11).Visible = False
End If
If 宠物已经购买(2) Then
起始按钮(12).Visible = True
Else
起始按钮(12).Visible = False
End If
If 宠物已经购买(3) Then
起始按钮(13).Visible = True
Else
起始按钮(13).Visible = False
End If
If 宠物已经购买(4) Then
起始按钮(14).Visible = True
Else
起始按钮(14).Visible = False
End If

End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'条件满足就升级
Dim sheng As Long
For sheng = 1 To 20 '防止一次生过多级必须连续五次单击
If 起始文字数组(2) >= 起始文字数组(1) * 10 + 30 Then
起始文字数组(2) = 起始文字数组(2) - (起始文字数组(1) * 10 + 30)
起始文字(2).Caption = 起始文字数组(2)
起始文字数组(1) = 起始文字数组(1) + 1
起始文字(1).Caption = 起始文字数组(1)
升级界面.Visible = True
End If
Next


If 正在游戏 = False Then
'隐藏玩家无法选择的按钮
If 起始文字数组(1) >= 5 Then
起始按钮(8).Visible = True
Else
起始按钮(8).Visible = False
End If
If 起始文字数组(1) >= 10 Then
起始按钮(6).Visible = True
Else
起始按钮(6).Visible = False
End If
If 起始文字数组(1) >= 15 Then
起始按钮(9).Visible = True
Else
起始按钮(9).Visible = False
End If
If 宠物已经购买(1) Then
起始按钮(11).Visible = True
Else
起始按钮(11).Visible = False
End If
If 宠物已经购买(2) Then
起始按钮(12).Visible = True
Else
起始按钮(12).Visible = False
End If
If 宠物已经购买(3) Then
起始按钮(13).Visible = True
Else
起始按钮(13).Visible = False
End If
If 宠物已经购买(4) Then
起始按钮(14).Visible = True
Else
起始按钮(14).Visible = False
End If
End If



'============================================
formmousex = X
formmousey = Y
'历史记录鼠标form还原：
If Y > 200 And Y < 560 And X > 6300 And X < 6660 Then
Else
起始按钮(1).Picture = 菜单图库(1).Picture
End If
'商店鼠标form还原：
If Y > 200 And Y < 560 And X > 7300 And X < 7660 Then
Else
起始按钮(2).Picture = 菜单图库(3).Picture
End If
'一个鼠标form还原：
If Y > 1500 And Y < 1860 And X > 900 And X < 2700 Then
Else
If 按键(1) <> 1 Then 起始按钮(3).Picture = 菜单图库(5).Picture

End If
'两个鼠标form还原：
If Y > 1500 And Y < 1860 And X > 3700 And X < 5900 Then
Else
If 按键(1) <> 2 Then 起始按钮(4).Picture = 菜单图库(7).Picture
End If
'简单鼠标form还原：
If Y > 2100 And Y < 1860 And X > 900 And X < 2700 Then
Else
If 按键(2) <> 1 Then 起始按钮(5).Picture = 菜单图库(9).Picture
End If
'一般鼠标form还原：
If Y > 2100 And Y < 1860 And X > 3700 And X < 5900 Then
Else
If 按键(2) <> 2 Then 起始按钮(6).Picture = 菜单图库(11).Picture
End If
'静心圣人鼠标form还原：
If Y > 2700 And Y < 3900 And X > 900 And X < 2100 Then
Else
If 按键(3) <> 1 Then 起始按钮(7).Picture = 菜单图库(13).Picture
End If
'能量领袖鼠标form还原：
If Y > 2700 And Y < 3900 And X > 3100 And X < 4300 Then
Else
If 按键(3) <> 2 Then 起始按钮(8).Picture = 菜单图库(17).Picture
End If
'变通公主鼠标form还原：
If Y > 2700 And Y < 3900 And X > 5300 And X < 6500 Then
Else
If 按键(3) <> 3 Then 起始按钮(9).Picture = 菜单图库(21).Picture
End If
'乌龟鼠标form还原：
If Y > 4340 And Y < 5540 And X > 900 And X < 2100 Then
Else
If 按键(4) <> 1 Then 起始按钮(10).Picture = 菜单图库(25).Picture
End If
'招财猫鼠标form还原：
If Y > 4340 And Y < 5540 And X > 3100 And X < 4300 Then
Else
If 按键(4) <> 2 Then 起始按钮(11).Picture = 菜单图库(27).Picture
End If
'虎不发威鼠标form还原：
If Y > 4340 And Y < 5540 And X > 5300 And X < 6500 Then
Else
If 按键(4) <> 3 Then 起始按钮(12).Picture = 菜单图库(29).Picture
End If
'学猴王鼠标form还原：
If Y > 5740 And Y < 6940 And X > 900 And X < 2100 Then
Else
If 按键(4) <> 4 Then 起始按钮(13).Picture = 菜单图库(31).Picture
End If
'精力十足鼠标form还原：
If Y > 5740 And Y < 6940 And X > 3100 And X < 4300 Then
Else
If 按键(4) <> 5 Then 起始按钮(14).Picture = 菜单图库(33).Picture
End If
'开始鼠标form还原：
If Y > 9300 And Y < 10200 And X > 2400 And X < 5100 Then
Else
If 开始按钮进行中 = False Then
起始按钮(15).Picture = 菜单图库(35).Picture
End If
End If
End Sub


Private Sub Timer1_Timer()
timing = True
 Dim o As Long, p As Long
 Dim k As Boolean
 Dim c(201) As Boolean
 '《消行与此行之上整体下降》
 For o = 1 To 20
   '看这行可不可以消：
   k = False
   For p = 1 To 10
   If Image1((o - 1) * 10 + p).Visible = False Then
   '记录不可消
   k = True
   End If
   Next
   '如果可以消
   If k = False Then '开始下移
       '奖励
   能量变量 = 能量变量 + 10
   能量数字.Caption = 能量变量
       '奖励图标
    Dim jl1 As Long
    For jl1 = 1 To 10
  With 奖励(奖励指针)
  .Picture = 游戏图库(31).Picture
  .Left = Image1((o - 1) * 10 + jl1).Left
  .Top = Image1((o - 1) * 10 + jl1).Top
  End With
  奖励剩余时间(奖励指针) = 50
    奖励指针 = 奖励指针 + 1
    If 奖励指针 > 50 Then 奖励指针 = 1
    Next
   
        '四方块下移
      Dim q As Long
      For q = 1 To 4
      b(q) = b(q) + 10
      Next
      '记录此行之上整体位置
      For q = 1 To 10 * (o - 1)
      If Image1(q).Visible = True Then
        c(q) = True
      Else
       c(q) = False
      End If
      Next
      '第一行变成无方块
      For q = 1 To 10
        Image1(q).Visible = False
      Next
      '把刚才记录的此行之上整体下移
      For q = 1 To 10 * (o - 1)
        If c(q) Then
        Image1(q + 10).Visible = True
        Else
        Image1(q + 10).Visible = False
        End If
      Next
   End If
Next
Dim 刚出新 As Boolean '《出新与结束》
刚出新 = False
If 出新 = True Then
 出新 = False
 Dim n As Long, i As Long
 Dim l As Boolean
 '本行确定新出方块样式：============
 n = (Int(Rnd * 11) + Second(Now)) Mod 11 + 1
 '‘’‘’'永恒之日控制开始：
 If 永恒之日变量 = 1 Then n = 2 + 3 * (Int(Rnd * 10) Mod 2) '此处编程不精
 If 永恒之日变量 = 2 Then n = 1
 If 永恒之日变量 = 3 Then n = 2
 If 永恒之日变量 = 4 Then n = 11
 '==================================
 
 bb = n 'vbkeyup使用
 b(1) = 6
 For i = 1 To 3
  b(i + 1) = 10 * ax(n, i) + 6 + ay(n, i)
 Next
 l = False
 For i = 1 To 4
 If Image1(b(i)).Visible = True Then
 '刚出就被挡
 l = True
 End If
 Next
 If l = True Then
'本局失败：~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
提示不同能力.Visible = False
timer1.Enabled = False
紧张时间系统.Enabled = False
游戏变脸时间系统.Enabled = False
树木皇后时间系统.Enabled = False
键盘不让用 = True
材料小偷时间系统.Enabled = False
能量小偷时间系统.Enabled = False
已经左右切换 = False
消失变量 = 11
消失时间系统.Enabled = True
消失时间系统.Interval = 300
'==============================
成功 = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 Else
  For i = 1 To 4
  Image1(b(i)).Visible = True
  Next
  刚出新 = True
 End If
End If
 ''''================================================
 If 刚出新 = False Then '《是否可下降》
    Dim f As Long '【下降】开始
    可下降 = True
    For f = 1 To 4
     Image1(b(f)).Visible = False '排除原图形的干扰
    Next
    For f = 1 To 4
     If (Image1(b(f) + 10).Visible = True) Or (b(f) + 10 > 200) Then '寻找不可下降条件是否成立
     可下降 = False
     End If
    Next
    If 可下降 Then
      For f = 1 To 4
        b(f) = b(f) + 10
        Image1(b(f)).Visible = True '正在下降
      Next
    Else
     For f = 1 To 4
      Image1(b(f)).Visible = True '“排除原图形的干扰”的还原
     Next
     出新 = True
    End If '【下降】结束
 End If
 ''''============================================
For f = 201 To 255
Image1(f).Visible = False
Next
Image1(0).Visible = False
timing = False
End Sub





Private Sub 材料小偷时间系统_Timer()
 Dim cl1 As Long, cl2 As Long, cl3 As Long 'cl1=i.cl2=最高点or行.cl3一行中放块数量
'寻找最高点：
cl2 = 201
For cl1 = 200 To 1 Step -1
If Image1(cl1).Visible = True Then cl2 = cl1
Next
'方块号变成行号
cl2 = (cl2 - 1) \ 10 + 1
'=======
For cl = 1 To 材料小偷变量
'寻找方块数大于五的最高行：
cl2 = cl2 - 1
Do While cl3 < 6 And cl2 <= 20
cl2 = cl2 + 1
cl3 = 0
For cl1 = (cl2 - 1) * 10 + 1 To cl2 * 10
If Image1(cl1).Visible = True Then cl3 = cl3 + 1
Next
Loop
'如果在20行以内就这行偷一个（可能偷到空白）
If cl2 <= 20 Then
Image1(cl2 * 10 - 9 + Int(Rnd * 10)).Visible = False
End If
Next
End Sub

Private Sub 虎不发威时间系统_Timer()
If 紧张时间 Mod 30 = 0 Then 虎威
End Sub







Private Sub 奖励时间系统_Timer()
Dim jls As Long
For jls = 1 To 50
奖励剩余时间(jls) = 奖励剩余时间(jls) - 1
If 奖励剩余时间(jls) <= 0 Then
奖励剩余时间(jls) = 0
奖励(jls).Visible = False
Else
  '闪烁奖励图标
  If 奖励剩余时间(jls) > 37 And 奖励剩余时间(jls) <= 50 Then
  奖励(jls).Visible = True
  End If
  If 奖励剩余时间(jls) > 26 And 奖励剩余时间(jls) <= 37 Then
  奖励(jls).Visible = False
  End If
  If 奖励剩余时间(jls) > 15 And 奖励剩余时间(jls) <= 26 Then
  奖励(jls).Visible = True
  End If
  '移动奖励
  If 奖励剩余时间(jls) Mod 3 = 0 And 奖励剩余时间(jls) <= 15 Then
  奖励(jls).Left = (奖励(jls).Left + 6000) / 2
  奖励(jls).Top = (奖励(jls).Top + 5800) / 2
  End If
End If
Next
End Sub

Private Sub 紧张时间系统_Timer()

If 紧张时间 = 1 Then
'成功：
提示不同能力.Visible = False
timer1.Enabled = False
紧张时间系统.Enabled = False
游戏变脸时间系统.Enabled = False
树木皇后时间系统.Enabled = False
键盘不让用 = True
材料小偷时间系统.Enabled = False
能量小偷时间系统.Enabled = False
已经左右切换 = False
消失变量 = 11
消失时间系统.Enabled = True
消失时间系统.Interval = 300
'==============================
成功 = True
End If





紧张时间 = 紧张时间 - 1
'游戏显示时间的控制：
游戏时间数字 = 紧张时间
'开始紧张：
If 紧张时间 = 96 Or 紧张时间 = 28 + 紧张延长 Then
timer1.Interval = 快的速度
正在加速 = True
End If
'结束紧张：
If 紧张时间 = 92 - 紧张延长 Or 紧张时间 = 24 Then
timer1.Interval = 慢的速度
正在加速 = False
End If
'防止 紧张时间过小超越接线：
If 紧张时间 < 0 Then 紧张时间 = 0
'=========================================================玩家电脑升级部分借用紧张时间系统=================================================
If 紧张时间 Mod 15 = 0 And 紧张时间 <> 0 And 紧张时间 <> 120 Then
游戏时间.Picture = 游戏图库(30).Picture
Dim qp As Long, qp1 As Long
qp1 = 201
For qp = 200 To 1 Step -1
If Image1(qp).Visible = True And qp <> b(1) And qp <> b(2) And qp <> b(3) And qp <> b(4) Then qp1 = qp
Next
If qp1 <= 100 Then
电脑等级变量 = 电脑等级变量 + 1
 If 电脑等级变量 = 1 Then 电脑等级.Picture = 游戏图库(26).Picture
 If 电脑等级变量 = 2 Then 电脑等级.Picture = 游戏图库(27).Picture
 If 电脑等级变量 = 3 Then 电脑等级.Picture = 游戏图库(28).Picture
 '主电脑升级能力变化：
 If 电脑等级变量 <= 3 Then
 If 电脑选人(1) = 1 Then 能量小偷力量 = 能量小偷力量 + 1
 If 电脑选人(1) = 2 Then 左右切换间隔 = 左右切换间隔 \ 2
 If 电脑选人(1) = 3 Then
快的速度 = 快的速度 - 7
慢的速度 = 慢的速度 - 50
End If
 If 电脑选人(1) = 4 Then 永恒之日变量 = 永恒之日变量 + 1
 If 电脑选人(1) = 5 Then 材料小偷变量 = 材料小偷变量 + 1
 If 电脑选人(1) = 6 Then 树木皇后变量 = 树木皇后变量 * 2 \ 3
 If 电脑选人(1) = 7 Then
If 电脑等级变量 = 1 Then
ax(4, 1) = 2
ax(4, 3) = 2
End If
If 电脑等级变量 = 2 Then ay(5, 2) = -2
If 电脑等级变量 = 3 Then
ay(7, 1) = -2
ay(7, 2) = 2
End If
End If
End If
'=====
Else
'=====
玩家等级变量 = 玩家等级变量 + 1
 If 玩家等级变量 = 1 Then 玩家等级.Picture = 游戏图库(26).Picture
 If 玩家等级变量 = 2 Then 玩家等级.Picture = 游戏图库(27).Picture
 If 玩家等级变量 = 3 Then 玩家等级.Picture = 游戏图库(28).Picture
 '玩家升级能力变化：
 If 玩家等级变量 <= 3 Then
 If 按键(3) = 1 Then
快的速度 = 快的速度 + 4
慢的速度 = 慢的速度 + 100
End If
 If 按键(3) = 2 Then
能量变量 = 能量变量 + 20 '升级相同
能量数字.Caption = 能量变量
End If
If 按键(3) = 3 Then '正斜反斜ok·+方直ok·余三不连?·全连?

  If 玩家等级变量 = 1 Then
li(10) = 9
li(8) = 11
li(6) = 2
li(3) = 6
  End If

 If 玩家等级变量 = 2 Then
li(2) = 3
li(3) = 6
li(6) = 8
li(8) = 9
li(9) = 10
li(10) = 11
li(11) = 2
 End If

 If 玩家等级变量 = 3 Then
li(7) = 2
li(11) = 1
 End If

End If
End If
 
 
End If
Else
游戏时间.Picture = 游戏图库(29).Picture

End If
'=========================================================================================================================================
End Sub

Private Sub 精力十足时间系统_Timer()
If 紧张时间 Mod 3 = 0 Then
能量变量 = 能量变量 + 1
能量数字.Caption = 能量变量
End If
End Sub

Private Sub 开始按钮动画时间系统_Timer()
If 开始按钮动画位置 = 1 Then
起始按钮(15).Picture = 菜单图库(38).Picture
End If
If 开始按钮动画位置 = 2 Then
起始按钮(15).Picture = 菜单图库(39).Picture
End If
If 开始按钮动画位置 = 3 Then
开始按钮进行中 = False
'=======================================开始按钮后半部分=============================
'菜单全部消失：
起始按钮(15).Visible = False
Dim si As Long
For si = 0 To 7
桌面星(si).Visible = False
流星(si).Visible = False
Next
For si = 0 To 15
桌面粒子(si).Visible = False
Next
'菜单按钮确定：（如果玩家没有选）
For si = 1 To 4
If 按键(si) = 0 Then 按键(si) = 1
Next
'宠物费用
If 按键(4) > 1 And 按键(4) <= 5 Then
起始文字数组(3) = 起始文字数组(3) - 3 - 按键(4)
End If

'游戏界面初始化：
游戏变脸时间系统.Enabled = True
If 起始文字数组(1) <= 8 Then 提示不同能力.Visible = True
For si = 1 To 200
Image1(si).Picture = 游戏图库(1).Picture
Next
If 起始文字数组(1) > 7 And 起始文字数组(1) <= 10 Then
Image1(101).Visible = True
Image1(104).Visible = True
Image1(107).Visible = True
Image1(110).Visible = True
Image1(111).Visible = True
Image1(120).Visible = True
Image1(121).Visible = True
Image1(125).Visible = True
Image1(126).Visible = True
Image1(130).Visible = True
Image1(131).Visible = True
Image1(132).Visible = True
Image1(139).Visible = True
Image1(140).Visible = True
Image1(141).Visible = True
Image1(142).Visible = True
Image1(145).Visible = True
Image1(146).Visible = True
Image1(149).Visible = True
Image1(150).Visible = True
Image1(151).Visible = True
Image1(152).Visible = True
Image1(159).Visible = True
Image1(160).Visible = True
Image1(161).Visible = True
Image1(162).Visible = True
Image1(163).Visible = True
Image1(165).Visible = True
Image1(166).Visible = True
Image1(168).Visible = True
Image1(169).Visible = True
Image1(170).Visible = True
Image1(171).Visible = True
Image1(172).Visible = True
Image1(173).Visible = True
Image1(178).Visible = True
Image1(179).Visible = True
Image1(180).Visible = True
Image1(181).Visible = True
Image1(182).Visible = True
Image1(183).Visible = True
Image1(188).Visible = True
Image1(189).Visible = True
Image1(190).Visible = True
Image1(191).Visible = True
Image1(192).Visible = True
Image1(193).Visible = True
Image1(194).Visible = True
Image1(197).Visible = True
Image1(198).Visible = True
Image1(199).Visible = True
Image1(200).Visible = True
End If
If 起始文字数组(1) > 10 Then
Image1(103).Visible = True
Image1(104).Visible = True
Image1(105).Visible = True
Image1(106).Visible = True
Image1(107).Visible = True
Image1(108).Visible = True
End If

'玩家二布局
With 玩家二
.Stretch = False
.Top = 9200
.Left = 6300
.Visible = True
End With
If 按键(4) = 1 Then
玩家二.Picture = 菜单图库(25).Picture
End If
If 按键(4) = 2 Then
玩家二.Picture = 菜单图库(27).Picture
End If
If 按键(4) = 3 Then
玩家二.Picture = 菜单图库(29).Picture
End If
If 按键(4) = 4 Then
玩家二.Picture = 菜单图库(31).Picture
End If
If 按键(4) = 5 Then
玩家二.Picture = 菜单图库(33).Picture
End If
'玩家一布局
With 玩家一
.Stretch = False
.Top = 7800
.Left = 6300
.Visible = True
End With
If 按键(3) = 1 Then
玩家一.Picture = 菜单图库(15).Picture
End If
If 按键(3) = 2 Then
玩家一.Picture = 菜单图库(19).Picture
End If
If 按键(3) = 3 Then
玩家一.Picture = 菜单图库(23).Picture
End If
'玩家名布局
With 玩家名
.Visible = True
.Top = 7400
.Left = 6500
.Height = 250
.Width = 800
.Caption = 人名(按键(3) + 7)
.ForeColor = &HFFFFFF
.BackColor = &H80000012
End With
'玩家等级布局：
With 玩家等级
.Visible = True
.Top = 7000
.Left = 6580
.Stretch = False
.Picture = 游戏图库(25).Picture
End With
'Q键布局
With q键
.Visible = True
.Top = 6400
.Left = 6300
.Stretch = False
.Picture = 游戏图库(32).Picture
End With
'Q键布局
With w键
.Visible = True
.Top = 6400
.Left = 6700
.Stretch = False
.Picture = 游戏图库(35).Picture
End With
'e键布局
With e键
.Visible = True
.Top = 6400
.Left = 7100
.Stretch = False
.Picture = 游戏图库(37).Picture
End With
'能量布局
With 能量
.Visible = True
.Top = 5800
.Left = 6200
.Stretch = False
.Picture = 游戏图库(31).Picture
End With
'能量数字布局
能量变量 = 10
With 能量数字
.Visible = True
.Top = 5900
.Left = 6600
.Height = 200
.Width = 900
.Caption = 能量变量
.ForeColor = &HFFFFFF
.BackColor = &H80000012
.Alignment = 2 '中间对齐
End With
'游戏时间布局
With 游戏时间
.Visible = True
.Top = 5200
.Left = 6400
.Stretch = False
.Picture = 游戏图库(29).Picture
End With
'游戏时间数字布局
With 游戏时间数字
.Visible = True
.Top = 5300
.Left = 6900
.Height = 200
.Width = 500
.Caption = 120
.Alignment = 2
.BackColor = &HFF00&
End With
'================================================================电脑选择人物：=============================================================
If 按键(1) = 2 Then
'等级高时选
电脑选人(1) = 1 + Int(Rnd * 7)
电脑选人(2) = 1 + Int(Rnd * 7)
Do While 电脑选人(1) = 电脑选人(2)
电脑选人(2) = 1 + Int(Rnd * 7)
Loop
'等级限制在最后
'等级一
If 起始文字数组(1) = 1 Then
电脑选人(1) = 1
电脑选人(2) = 3
End If
'等级2
If 起始文字数组(1) = 2 Then
电脑选人(1) = 3
电脑选人(2) = 4
End If
'等级3
If 起始文字数组(1) = 3 Then
电脑选人(1) = 4
电脑选人(2) = 5
End If
'等级4
If 起始文字数组(1) = 4 Then
电脑选人(1) = 5
电脑选人(2) = 2
End If
'等级5
If 起始文字数组(1) = 5 Then
电脑选人(1) = 2
电脑选人(2) = 7
End If
'等级6
If 起始文字数组(1) = 6 Then
电脑选人(1) = 7
电脑选人(2) = 6
End If
'等级7
If 起始文字数组(1) = 7 Then
电脑选人(1) = 6
电脑选人(2) = 7
End If
Else
'等级高时选
电脑选人(1) = 1 + Int(7 * Rnd)
电脑选人(2) = 0
'等级限制在最后
'等级一
If 起始文字数组(1) = 1 Then
电脑选人(1) = 1
电脑选人(2) = 0
End If
'等级2
If 起始文字数组(1) = 2 Then
电脑选人(1) = 3
电脑选人(2) = 0
End If
'等级3
If 起始文字数组(1) = 3 Then
电脑选人(1) = 4
电脑选人(2) = 0
End If
'等级4
If 起始文字数组(1) = 4 Then
电脑选人(1) = 5
电脑选人(2) = 0
End If
'等级5
If 起始文字数组(1) = 5 Then
电脑选人(1) = 2
电脑选人(2) = 0
End If
'等级6
If 起始文字数组(1) = 6 Then
电脑选人(1) = 7
电脑选人(2) = 0
End If
'等级7
If 起始文字数组(1) = 7 Then
电脑选人(1) = 6
电脑选人(2) = 0
End If
End If

'电脑选人王帅控制============================================================================必修改
'电脑选人(1) = 6
'电脑选人(2) = 0
'电脑一布局：
With 电脑一
.Visible = True
.Top = 1700
.Left = 6300
.Stretch = False
End With
If 电脑选人(1) = 1 Then 电脑一.Picture = 游戏图库(11).Picture
If 电脑选人(1) = 2 Then 电脑一.Picture = 游戏图库(13).Picture
If 电脑选人(1) = 3 Then 电脑一.Picture = 游戏图库(15).Picture
If 电脑选人(1) = 4 Then 电脑一.Picture = 游戏图库(17).Picture
If 电脑选人(1) = 5 Then 电脑一.Picture = 游戏图库(19).Picture
If 电脑选人(1) = 6 Then 电脑一.Picture = 游戏图库(21).Picture
If 电脑选人(1) = 7 Then 电脑一.Picture = 游戏图库(23).Picture
'电脑二布局：
If 电脑选人(2) <> 0 Then
With 电脑二
.Visible = True
.Top = 300
.Left = 6300
.Stretch = False
End With
If 电脑选人(2) = 1 Then 电脑二.Picture = 游戏图库(11).Picture
If 电脑选人(2) = 2 Then 电脑二.Picture = 游戏图库(13).Picture
If 电脑选人(2) = 3 Then 电脑二.Picture = 游戏图库(15).Picture
If 电脑选人(2) = 4 Then 电脑二.Picture = 游戏图库(17).Picture
If 电脑选人(2) = 5 Then 电脑二.Picture = 游戏图库(19).Picture
If 电脑选人(2) = 6 Then 电脑二.Picture = 游戏图库(21).Picture
If 电脑选人(2) = 7 Then 电脑二.Picture = 游戏图库(23).Picture
End If
'电脑一名布局
With 电脑名
.Visible = True
.Top = 3200
.Left = 6200
.Height = 250
.Width = 1500
.Caption = " 敌人：" & 人名(电脑选人(1))
.ForeColor = &HFFFFFF
.BackColor = &H80000012
End With
'电脑一等级布局：
With 电脑等级
.Visible = True
.Top = 3700
.Left = 6580
.Stretch = False
.Picture = 游戏图库(25).Picture
End With
'紧张系统开始起作用
紧张时间系统.Enabled = True
紧张时间 = 120
正在加速 = False
快的速度 = 168 '初始168,范围140-188,电脑-7,玩家or宠物+4
慢的速度 = 1000 '初始1000，范围800-1500，电脑-50，玩家or宠物+100
'电脑人物功能载入：
'''''''能量小偷载入
能量小偷时间系统.Enabled = False
If 电脑选人(1) = 1 Or 电脑选人(2) = 1 Then
能量小偷时间系统.Enabled = True
能量小偷时间系统.Interval = 2000
能量小偷力量 = 1
End If
'''''''左右不分载入
左右时间系统.Enabled = False
If 电脑选人(1) = 2 Or 电脑选人(2) = 2 Then
左右切换间隔 = 11 '1,2,5
左右时间系统.Enabled = True
左右时间系统.Interval = 1000
已经左右切换 = True
End If
'''''''流星之王载入
If 电脑选人(1) = 3 Or 电脑选人(2) = 3 Then
快的速度 = 快的速度 - 7
慢的速度 = 慢的速度 - 50
End If
'''''''永恒之日载入
 永恒之日变量 = 0 '三直-三-直-斜
If 电脑选人(1) = 4 Or 电脑选人(2) = 4 Then
永恒之日变量 = 1
End If
'‘’‘’‘’材料小偷载入
材料小偷时间系统.Enabled = False
材料小偷变量 = 0
材料小偷时间系统.Interval = 3000
If 电脑选人(1) = 5 Or 电脑选人(2) = 5 Then
材料小偷时间系统.Enabled = True
材料小偷变量 = 1
End If
'’‘’‘’‘’‘’‘’树后载入能力
树木皇后时间系统.Enabled = False
If 电脑选人(1) = 6 Or 电脑选人(2) = 6 Then
树木皇后时间系统.Enabled = True
树木皇后时间系统.Interval = 1000
树木皇后变量 = 30 '30-20-13-8
End If
'''''''''''''''''变异蜗牛载入
ay(1, 1) = 0
ax(4, 1) = 1
ay(4, 1) = -1
ax(4, 2) = 1
ay(4, 2) = 0
ax(4, 3) = 1
ay(4, 3) = 1
ax(5, 1) = 1
ay(5, 1) = -1
ax(5, 2) = 1
ay(5, 2) = 0
ax(5, 3) = 2
ay(5, 3) = 0
ax(7, 1) = 0
ay(7, 1) = -1
ax(7, 2) = 1
ay(7, 2) = 0
ax(7, 3) = 0
ay(7, 3) = 1
If 电脑选人(1) = 7 Or 电脑选人(2) = 7 Then
ay(1, 1) = 2
End If
'中等载入
If 按键(2) = 2 Then
快的速度 = 快的速度 - 25
慢的速度 = 慢的速度 - 500
End If
'玩家人物功能载入
''''''''''静心圣人载入
If 按键(3) = 1 Then
快的速度 = 快的速度 + 4
慢的速度 = 慢的速度 + 100
End If
''''''''''能量领袖载入
If 按键(3) = 2 Then
能量变量 = 能量变量 + 20 '升级相同
能量数字.Caption = 能量变量
End If
'''''''''变通公主载入
li(1) = 4
li(4) = 5
li(5) = 7
li(7) = 1
li(2) = 3
li(3) = 2
li(6) = 6
li(8) = 9
li(9) = 8
li(10) = 11
li(11) = 10
If 按键(3) = 3 Then '正斜反斜ok·+方直ok·余三不连?·全连?
li(10) = 9
li(8) = 11
End If
'宠物功能载入
''''''''乌龟功能载入
If 按键(4) = 1 Then
快的速度 = 快的速度 + 4
慢的速度 = 慢的速度 + 100
End If
''''''''招财猫功能载入
If 按键(4) = 2 Then















End If
''''''''虎不发威功能载入
虎不发威时间系统.Enabled = False
If 按键(3) = 3 Then
虎不发威时间系统.Enabled = True
虎不发威时间系统.Interval = 1000
End If
''''''''学猴王载入
If 按键(4) = 4 Then













End If
''''''''精力十足载入
精力十足时间系统.Enabled = False
If 按键(4) = 5 Then
精力十足时间系统.Interval = 1000
精力十足时间系统.Enabled = True
End If
'等级变量初始化
玩家等级变量 = 0
电脑等级变量 = 0
'第一次速度设定（需要靠后）
timer1.Interval = 慢的速度
'下落开始：（需要靠后）
timer1.Enabled = True
键盘不让用 = False

'====================================================================================
End If
开始按钮动画位置 = 开始按钮动画位置 + 1
If 开始按钮动画位置 > 4 Then
开始按钮动画位置 = 5 '不知是否可以终止timer
End If
End Sub

Private Sub 流星时间系统_Timer()
Dim k1 As Long
For k1 = 1 To 7
If 流星(k1).Top > 11000 Then
流星(k1).Top = -1000
End If
流星(k1).Top = 流星(k1).Top + 1500
Next

End Sub

Private Sub 买虎不发威_Click()
If 宠物已经购买(2) = False And 起始文字数组(3) >= 60 Then
宠物已经购买(2) = True
起始文字数组(3) = 起始文字数组(3) - 60
起始文字(3).Caption = 起始文字数组(3)
End If
买招财猫.Visible = False
买虎不发威.Visible = False
买学猴王.Visible = False
买精力十足.Visible = False
商店.Visible = False
End Sub

Private Sub 买精力十足_Click()
If 宠物已经购买(4) = False And 起始文字数组(3) >= 80 Then
宠物已经购买(4) = True
起始文字数组(3) = 起始文字数组(3) - 80
起始文字(3).Caption = 起始文字数组(3)
End If
买招财猫.Visible = False
买虎不发威.Visible = False
买学猴王.Visible = False
买精力十足.Visible = False
商店.Visible = False
End Sub

Private Sub 买学猴王_Click()
If 宠物已经购买(3) = False And 起始文字数组(3) >= 70 Then
宠物已经购买(3) = True
起始文字数组(3) = 起始文字数组(3) - 70
起始文字(3).Caption = 起始文字数组(3)
End If
买招财猫.Visible = False
买虎不发威.Visible = False
买学猴王.Visible = False
买精力十足.Visible = False
商店.Visible = False
End Sub

Private Sub 买招财猫_Click()
If 宠物已经购买(1) = False And 起始文字数组(3) >= 50 Then
宠物已经购买(1) = True
起始文字数组(3) = 起始文字数组(3) - 50
起始文字(3).Caption = 起始文字数组(3)
End If
买招财猫.Visible = False
买虎不发威.Visible = False
买学猴王.Visible = False
买精力十足.Visible = False
商店.Visible = False
End Sub

Private Sub 能量小偷时间系统_Timer()
If 能量变量 > 能量小偷力量 Then 能量变量 = 能量变量 - 能量小偷力量
能量数字.Caption = 能量变量

End Sub

Private Sub 起始按钮_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If formmousey > 100 And formmousey < 660 And formmousex > 6100 And formmousex < 6860 Then '扩大100/高宽/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(1).Picture = 菜单图库(2).Picture
End If
If formmousey > 100 And formmousey < 660 And formmousex > 7100 And formmousex < 7860 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(2).Picture = 菜单图库(4).Picture
End If
If formmousey > 1400 And formmousey < 1960 And formmousex > 700 And formmousex < 2900 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(3).Picture = 菜单图库(6).Picture
End If
If formmousey > 1400 And formmousey < 1960 And formmousex > 3500 And formmousex < 6100 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(4).Picture = 菜单图库(8).Picture
End If
If formmousey > 2000 And formmousey < 2560 And formmousex > 700 And formmousex < 2900 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(5).Picture = 菜单图库(10).Picture
End If
If formmousey > 2000 And formmousey < 2560 And formmousex > 3500 And formmousex < 6100 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(6).Picture = 菜单图库(12).Picture
End If
If formmousey > 2600 And formmousey < 4000 And formmousex > 700 And formmousex < 2300 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(7).Picture = 菜单图库(14).Picture
End If
If formmousey > 2600 And formmousey < 4000 And formmousex > 2900 And formmousex < 4700 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(8).Picture = 菜单图库(18).Picture
End If
If formmousey > 2600 And formmousey < 4000 And formmousex > 5100 And formmousex < 6700 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(9).Picture = 菜单图库(22).Picture
End If
If formmousey > 4240 And formmousey < 5640 And formmousex > 700 And formmousex < 2300 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(10).Picture = 菜单图库(26).Picture
End If
If formmousey > 4240 And formmousey < 5640 And formmousex > 2900 And formmousex < 4700 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(11).Picture = 菜单图库(28).Picture
End If
If formmousey > 4240 And formmousey < 5640 And formmousex > 5100 And formmousex < 6700 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(12).Picture = 菜单图库(30).Picture
End If
If formmousey > 5640 And formmousey < 7040 And formmousex > 700 And formmousex < 2300 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(13).Picture = 菜单图库(32).Picture
End If
If formmousey > 5640 And formmousey < 7040 And formmousex > 2900 And formmousex < 4700 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(14).Picture = 菜单图库(34).Picture
End If
If formmousey > 9200 And formmousey < 10300 And formmousex > 2200 And formmousex < 5300 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
If 开始按钮进行中 = False Then 起始按钮(15).Picture = 菜单图库(36).Picture
End If
End Sub



Private Sub 起始按钮_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'历史记录按键鼠标抬起：
If formmousey > 100 And formmousey < 660 And formmousex > 6100 And formmousex < 6860 Then '扩大100/高宽/200个单位来验证form mouse 位置,来确定控件数组小控件

End If
'商店按键鼠标抬起：
If formmousey > 100 And formmousey < 660 And formmousex > 7100 And formmousex < 7860 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
买招财猫.Visible = True
买虎不发威.Visible = True
买学猴王.Visible = True
买精力十足.Visible = True
商店.Visible = True
End If
'一个按键鼠标抬起：
If formmousey > 1400 And formmousey < 1960 And formmousex > 700 And formmousex < 2900 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(3).Picture = 菜单图库(6).Picture
按键(1) = 1
End If
'两个按键鼠标抬起：
If formmousey > 1400 And formmousey < 1960 And formmousex > 3500 And formmousex < 6100 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(4).Picture = 菜单图库(8).Picture
按键(1) = 2
End If
'简单按键鼠标抬起：
If formmousey > 2000 And formmousey < 2560 And formmousex > 700 And formmousex < 2900 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(5).Picture = 菜单图库(10).Picture
按键(2) = 1
End If
'一般按键鼠标抬起：
If formmousey > 2000 And formmousey < 2560 And formmousex > 3500 And formmousex < 6100 And 起始文字数组(1) >= 10 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(6).Picture = 菜单图库(12).Picture
按键(2) = 2
End If
'静心圣人按键鼠标抬起：
If formmousey > 2600 And formmousey < 4000 And formmousex > 700 And formmousex < 2300 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(7).Picture = 菜单图库(14).Picture
按键(3) = 1
End If
'能量领袖按键鼠标抬起：
If formmousey > 2600 And formmousey < 4000 And formmousex > 2900 And formmousex < 4700 And 起始文字数组(1) >= 5 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(8).Picture = 菜单图库(18).Picture
按键(3) = 2
End If
'变通公主按键鼠标抬起：
If formmousey > 2600 And formmousey < 4000 And formmousex > 5100 And formmousex < 6700 And 起始文字数组(1) >= 15 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(9).Picture = 菜单图库(22).Picture
按键(3) = 3
End If
'乌龟按键鼠标抬起：
If formmousey > 4240 And formmousey < 5640 And formmousex > 700 And formmousex < 2300 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(10).Picture = 菜单图库(26).Picture
按键(4) = 1
End If
'招财猫按键鼠标抬起：
If formmousey > 4240 And formmousey < 5640 And formmousex > 2900 And formmousex < 4700 And 宠物已经购买(1) Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(11).Picture = 菜单图库(28).Picture
按键(4) = 2
End If
'虎不发威按键鼠标抬起：
If formmousey > 4240 And formmousey < 5640 And formmousex > 5100 And formmousex < 6700 And 宠物已经购买(2) Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(12).Picture = 菜单图库(30).Picture
按键(4) = 3
End If
'学猴王按键鼠标抬起：
If formmousey > 5640 And formmousey < 7040 And formmousex > 700 And formmousex < 2300 And 宠物已经购买(3) Then  '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(13).Picture = 菜单图库(32).Picture
按键(4) = 4
End If
'精力十足按键鼠标抬起：
If formmousey > 5640 And formmousey < 7040 And formmousex > 2900 And formmousex < 4700 And 宠物已经购买(4) Then  '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
起始按钮(14).Picture = 菜单图库(34).Picture
按键(4) = 5
End If
'开始按键鼠标抬起：
'====================================开始按钮前半部分===========================================
If formmousey > 9200 And formmousey < 10300 And formmousex > 2200 And formmousex < 5300 Then '扩大100/200个单位来验证form mouse 位置,来确定控件数组小控件
正在游戏 = True
 '菜单画面主要图片消失：
 Dim kf As Long
 For kf = 0 To 23
 起始按钮(kf).Visible = False
 Next
 For kf = 0 To 3
 起始文字(kf).Visible = False
 Next
 桌面树.Visible = False
 '开始按钮动画：
 起始按钮(15).Visible = True
 开始按钮动画位置 = 1
 开始按钮进行中 = True
 起始按钮(15).Picture = 菜单图库(37).Picture
 开始按钮动画时间系统.Enabled = True
'===============================================================================================
End If
End Sub

Private Sub 商店_Click()
买招财猫.Visible = False
买虎不发威.Visible = False
买学猴王.Visible = False
买精力十足.Visible = False
商店.Visible = False

End Sub

Private Sub 升级界面_Click()
升级界面.Visible = False
End Sub

Private Sub 树木皇后时间系统_Timer()
If 紧张时间 Mod 树木皇后变量 = 0 And 紧张时间 < 110 Then
Dim sh As Long
For sh = 1 To 200
Image1(sh).Visible = True
Next
For sh = 1 To 40
Image1(sh).Visible = False
Next
Image1(41).Visible = False
Image1(43).Visible = False
Image1(44).Visible = False
Image1(45).Visible = False
Image1(46).Visible = False
Image1(47).Visible = False
Image1(48).Visible = False
Image1(50).Visible = False
Image1(51).Visible = False
Image1(52).Visible = False
Image1(54).Visible = False
Image1(55).Visible = False
Image1(56).Visible = False
Image1(57).Visible = False
Image1(59).Visible = False
Image1(60).Visible = False
Image1(62).Visible = False
Image1(64).Visible = False
Image1(65).Visible = False
Image1(66).Visible = False
Image1(67).Visible = False
Image1(69).Visible = False
Image1(71).Visible = False
Image1(73).Visible = False
Image1(78).Visible = False
Image1(80).Visible = False
Image1(82).Visible = False
Image1(83).Visible = False
Image1(84).Visible = False
Image1(87).Visible = False
Image1(88).Visible = False
Image1(89).Visible = False
Image1(91).Visible = False
Image1(100).Visible = False
Image1(102).Visible = False
Image1(103).Visible = False
Image1(108).Visible = False
Image1(109).Visible = False
Image1(111).Visible = False
Image1(113).Visible = False
Image1(114).Visible = False
Image1(117).Visible = False
Image1(118).Visible = False
Image1(120).Visible = False
For sh = 13 To 19
Image1((sh - 1) * 10 + 1).Visible = False
Image1((sh - 1) * 10 + 2).Visible = False
Image1((sh - 1) * 10 + 3).Visible = False
Image1((sh - 1) * 10 + 4).Visible = False
Image1((sh - 1) * 10 + 7).Visible = False
Image1((sh - 1) * 10 + 8).Visible = False
Image1((sh - 1) * 10 + 9).Visible = False
Image1((sh - 1) * 10 + 10).Visible = False
Next
Image1(191).Visible = False
Image1(192).Visible = False
Image1(193).Visible = False
Image1(198).Visible = False
Image1(199).Visible = False
Image1(200).Visible = False
End If
End Sub

Private Sub 消失时间系统_Timer()
Dim xs As Long
If 消失变量 = 11 Then
For xs = 1 To 200
Image1(xs).Picture = 游戏图库(10).Picture
Next
End If
If 消失变量 < 11 And 消失变量 >= 1 Then
For xs = 1 To 200
Image1(xs).Picture = 游戏图库(37 + 12 - 消失变量).Picture
Next
End If
If 消失变量 = 0 Then
'============================奖励界面===============================
'出现部分：
If 成功 Then
结果(0).Visible = True
结果(0).Left = 3000
结果(0).Top = 3000
结果(0).Stretch = False
Else
结果(1).Visible = True
结果(1).Left = 3000
结果(1).Top = 3000
结果(1).Stretch = False
End If

经验增量 = 120 - 紧张时间
金钱增量 = 15 + 玩家等级变量 - 电脑等级变量
If 按键(1) = 2 Then
经验增量 = 2 * 经验增量
金钱增量 = 金钱增量 * 2
End If
If 按键(2) = 2 Then
经验增量 = 2 * 经验增量
金钱增量 = 金钱增量 * 2
End If
If 成功 Then
经验增量 = 2 * 经验增量
金钱增量 = 金钱增量 * 2
End If
If 按键(4) = 2 Then 金钱增量 = 金钱增量 * 2
If 按键(4) = 4 Then 经验增量 = 2 * 经验增量

起始按钮(17).Visible = True
起始按钮(17).Left = 3000
起始按钮(17).Top = 6000
起始文字(2).Visible = True
起始文字(2).Left = 3500
起始文字(2).Top = 6000
起始文字(2).Caption = "+" & 经验增量
起始按钮(18).Visible = True
起始按钮(18).Left = 3000
起始按钮(18).Top = 6500
起始文字(3).Visible = True
起始文字(3).Left = 3500
起始文字(3).Top = 6500
起始文字(3).Caption = "+" & 金钱增量
'消失部分：
Dim xs1 As Long
For xs1 = 1 To 200
Image1(xs1).Visible = False
Next
玩家一.Visible = False
玩家二.Visible = False
电脑一.Visible = False
电脑二.Visible = False
玩家名.Visible = False
电脑名.Visible = False
玩家等级.Visible = False
电脑等级.Visible = False
游戏时间.Visible = False
游戏时间数字.Visible = False
能量.Visible = False
能量数字.Visible = False
q键.Visible = False
w键.Visible = False
e键.Visible = False
'其他部分：
该回菜单 = True
'===================================================================
End If
消失变量 = 消失变量 - 1
If 消失变量 < -100 Then 消失变量 = -1
End Sub

Private Sub 游戏变脸时间系统_Timer()
变脸时间 = 变脸时间 + 1
If 变脸时间 > 100 Then 变脸时间 = 0
Dim bian As Long, btop As Long, sb As Long
'寻找最高点方块号码：
btop = 201
For bian = 200 To 1 Step -1
If Image1(bian).Visible = True And bian <> b(1) And bian <> b(2) And bian <> b(3) And bian <> b(4) Then
btop = bian
End If
Next
'开始变脸：
If 正在加速 Then
'闪电脸部分：
 If 变脸时间 Mod 3 = 0 Then
 For sb = 1 To 200
 Image1(sb).Picture = 游戏图库(7).Picture
 Next
Else
 For sb = 1 To 200
 Image1(sb).Picture = 游戏图库(8).Picture
 Next
End If
Else
If btop <= 100 Then
'哭脸部分：
If 变脸时间 Mod 3 = 0 Then
 For sb = 1 To 200
 Image1(sb).Picture = 游戏图库(4).Picture
 Next
Else
 For sb = 1 To 200
 Image1(sb).Picture = 游戏图库(3).Picture
 Next
End If
Else
'笑脸部分：
If 变脸时间 Mod 9 = 0 Then
 Image1(151 + Int(Rnd * 50)).Picture = 游戏图库(9).Picture
 Image1(151 + Int(Rnd * 50)).Picture = 游戏图库(9).Picture
 Image1(151 + Int(Rnd * 50)).Picture = 游戏图库(9).Picture
 Image1(151 + Int(Rnd * 50)).Picture = 游戏图库(9).Picture
 Image1(151 + Int(Rnd * 50)).Picture = 游戏图库(9).Picture
 Image1(151 + Int(Rnd * 50)).Picture = 游戏图库(9).Picture
Else
If 变脸时间 Mod 3 = 0 Then
 For sb = 1 To 200
 Image1(sb).Picture = 游戏图库(2).Picture
 Next
Else
 For sb = 1 To 200
 Image1(sb).Picture = 游戏图库(1).Picture
 Next
End If
End If
End If
End If
End Sub

Private Sub 桌面粒子时间系统_Timer()
Dim k1 As Long
For k1 = 1 To 15
If 桌面粒子(k1).Top < 0 Or 桌面粒子(k6).Left < 0 Then '超过界限则令其回到右下角
   桌面粒子(k1).Top = 11700
   桌面粒子(k1).Left = 9100
   桌面粒子(k1).Picture = 菜单图库(51).Picture
End If
If 桌面粒子(k1).Top < 3500 Or 桌面粒子(k1).Left < 3000 Then '左上角黄色
桌面粒子(k1).Picture = 菜单图库(53).Picture
Else
   If 桌面粒子(k1).Top > 7000 And 桌面粒子(k1).Left > 6000 Then '右下角橙色
   桌面粒子(k1).Picture = 菜单图库(51).Picture
   Else
   桌面粒子(k1).Picture = 菜单图库(52).Picture '中间部分中间色
   End If
End If
桌面粒子(k1).Top = 桌面粒子(k1).Top - Int(Rnd * 20) '向左上移动
桌面粒子(k1).Left = 桌面粒子(k1).Left - Int(Rnd * 20)
Next

End Sub



Private Sub 桌面树_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If 桌面树果实 = True Then
起始文字数组(3) = 起始文字数组(3) + 1
起始文字(3).Caption = 起始文字数组(3)
桌面树果实 = False
桌面树.Picture = 菜单图库(56).Picture

       '奖励图标
  
  With 奖励(奖励指针)
  .Picture = 菜单图库(42).Picture
  .Left = 桌面树.Left + 300
  .Top = 桌面树.Top + 300
  End With
  奖励剩余时间(奖励指针) = 50
    奖励指针 = 奖励指针 + 1
    If 奖励指针 > 50 Then 奖励指针 = 1
End If
End Sub

Private Sub 桌面树时间系统_Timer()
桌面树时间 = 桌面树时间 + 1
If 桌面树时间 > 4 Then
桌面树时间 = 0
桌面树果实 = True
End If
If 桌面树果实 Then
If Rnd > 0.5 Then
桌面树.Picture = 菜单图库(54).Picture
Else
桌面树.Picture = 菜单图库(55).Picture
End If
End If
End Sub

Private Sub 桌面星时间系统_Timer()
Dim k1 As Long
For k1 = 1 To 6
桌面星(k1).Picture = 菜单图库(47 + Int(Rnd * 4)).Picture
Next
End Sub

Private Sub 左右时间系统_Timer()
If (紧张时间 - 1) Mod 左右切换间隔 = 0 Then
If 已经左右切换 Then
已经左右切换 = False
Else
已经左右切换 = True
End If
End If

End Sub
