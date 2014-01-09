VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "扫雷"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7695
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton btnStart 
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   639
      Left            =   7440
      TabIndex        =   640
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   638
      Left            =   7200
      TabIndex        =   639
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   637
      Left            =   6960
      TabIndex        =   638
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   636
      Left            =   6720
      TabIndex        =   637
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   635
      Left            =   6480
      TabIndex        =   636
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   634
      Left            =   6240
      TabIndex        =   635
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   633
      Left            =   6000
      TabIndex        =   634
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   632
      Left            =   5760
      TabIndex        =   633
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   631
      Left            =   5520
      TabIndex        =   632
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   630
      Left            =   5280
      TabIndex        =   631
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   629
      Left            =   5040
      TabIndex        =   630
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   628
      Left            =   4800
      TabIndex        =   629
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   627
      Left            =   4560
      TabIndex        =   628
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   626
      Left            =   4320
      TabIndex        =   627
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   625
      Left            =   4080
      TabIndex        =   626
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   624
      Left            =   3840
      TabIndex        =   625
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   623
      Left            =   3600
      TabIndex        =   624
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   622
      Left            =   3360
      TabIndex        =   623
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   621
      Left            =   3120
      TabIndex        =   622
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   620
      Left            =   2880
      TabIndex        =   621
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   619
      Left            =   2640
      TabIndex        =   620
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   618
      Left            =   2400
      TabIndex        =   619
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   617
      Left            =   2160
      TabIndex        =   618
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   616
      Left            =   1920
      TabIndex        =   617
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   615
      Left            =   1680
      TabIndex        =   616
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   614
      Left            =   1440
      TabIndex        =   615
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   613
      Left            =   1200
      TabIndex        =   614
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   612
      Left            =   960
      TabIndex        =   613
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   611
      Left            =   720
      TabIndex        =   612
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   610
      Left            =   480
      TabIndex        =   611
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   609
      Left            =   240
      TabIndex        =   610
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   608
      Left            =   0
      TabIndex        =   609
      Top             =   5520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   607
      Left            =   7440
      TabIndex        =   608
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   606
      Left            =   7200
      TabIndex        =   607
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   605
      Left            =   6960
      TabIndex        =   606
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   604
      Left            =   6720
      TabIndex        =   605
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   603
      Left            =   6480
      TabIndex        =   604
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   602
      Left            =   6240
      TabIndex        =   603
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   601
      Left            =   6000
      TabIndex        =   602
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   600
      Left            =   5760
      TabIndex        =   601
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   599
      Left            =   5520
      TabIndex        =   600
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   598
      Left            =   5280
      TabIndex        =   599
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   597
      Left            =   5040
      TabIndex        =   598
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   596
      Left            =   4800
      TabIndex        =   597
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   595
      Left            =   4560
      TabIndex        =   596
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   594
      Left            =   4320
      TabIndex        =   595
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   593
      Left            =   4080
      TabIndex        =   594
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   592
      Left            =   3840
      TabIndex        =   593
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   591
      Left            =   3600
      TabIndex        =   592
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   590
      Left            =   3360
      TabIndex        =   591
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   589
      Left            =   3120
      TabIndex        =   590
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   588
      Left            =   2880
      TabIndex        =   589
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   587
      Left            =   2640
      TabIndex        =   588
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   586
      Left            =   2400
      TabIndex        =   587
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   585
      Left            =   2160
      TabIndex        =   586
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   584
      Left            =   1920
      TabIndex        =   585
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   583
      Left            =   1680
      TabIndex        =   584
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   582
      Left            =   1440
      TabIndex        =   583
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   581
      Left            =   1200
      TabIndex        =   582
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   580
      Left            =   960
      TabIndex        =   581
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   579
      Left            =   720
      TabIndex        =   580
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   578
      Left            =   480
      TabIndex        =   579
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   577
      Left            =   240
      TabIndex        =   578
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   576
      Left            =   0
      TabIndex        =   577
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   575
      Left            =   7440
      TabIndex        =   576
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   574
      Left            =   7200
      TabIndex        =   575
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   573
      Left            =   6960
      TabIndex        =   574
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   572
      Left            =   6720
      TabIndex        =   573
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   571
      Left            =   6480
      TabIndex        =   572
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   570
      Left            =   6240
      TabIndex        =   571
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   569
      Left            =   6000
      TabIndex        =   570
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   568
      Left            =   5760
      TabIndex        =   569
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   567
      Left            =   5520
      TabIndex        =   568
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   566
      Left            =   5280
      TabIndex        =   567
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   565
      Left            =   5040
      TabIndex        =   566
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   564
      Left            =   4800
      TabIndex        =   565
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   563
      Left            =   4560
      TabIndex        =   564
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   562
      Left            =   4320
      TabIndex        =   563
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   561
      Left            =   4080
      TabIndex        =   562
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   560
      Left            =   3840
      TabIndex        =   561
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   559
      Left            =   3600
      TabIndex        =   560
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   558
      Left            =   3360
      TabIndex        =   559
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   557
      Left            =   3120
      TabIndex        =   558
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   556
      Left            =   2880
      TabIndex        =   557
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   555
      Left            =   2640
      TabIndex        =   556
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   554
      Left            =   2400
      TabIndex        =   555
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   553
      Left            =   2160
      TabIndex        =   554
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   552
      Left            =   1920
      TabIndex        =   553
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   551
      Left            =   1680
      TabIndex        =   552
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   550
      Left            =   1440
      TabIndex        =   551
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   549
      Left            =   1200
      TabIndex        =   550
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   548
      Left            =   960
      TabIndex        =   549
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   547
      Left            =   720
      TabIndex        =   548
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   546
      Left            =   480
      TabIndex        =   547
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   545
      Left            =   240
      TabIndex        =   546
      Top             =   5040
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   544
      Left            =   0
      TabIndex        =   545
      Top             =   5040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   543
      Left            =   7440
      TabIndex        =   544
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   542
      Left            =   7200
      TabIndex        =   543
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   541
      Left            =   6960
      TabIndex        =   542
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   540
      Left            =   6720
      TabIndex        =   541
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   539
      Left            =   6480
      TabIndex        =   540
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   538
      Left            =   6240
      TabIndex        =   539
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   537
      Left            =   6000
      TabIndex        =   538
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   536
      Left            =   5760
      TabIndex        =   537
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   535
      Left            =   5520
      TabIndex        =   536
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   534
      Left            =   5280
      TabIndex        =   535
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   533
      Left            =   5040
      TabIndex        =   534
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   532
      Left            =   4800
      TabIndex        =   533
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   531
      Left            =   4560
      TabIndex        =   532
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   530
      Left            =   4320
      TabIndex        =   531
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   529
      Left            =   4080
      TabIndex        =   530
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   528
      Left            =   3840
      TabIndex        =   529
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   527
      Left            =   3600
      TabIndex        =   528
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   526
      Left            =   3360
      TabIndex        =   527
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   525
      Left            =   3120
      TabIndex        =   526
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   524
      Left            =   2880
      TabIndex        =   525
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   523
      Left            =   2640
      TabIndex        =   524
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   522
      Left            =   2400
      TabIndex        =   523
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   521
      Left            =   2160
      TabIndex        =   522
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   520
      Left            =   1920
      TabIndex        =   521
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   519
      Left            =   1680
      TabIndex        =   520
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   518
      Left            =   1440
      TabIndex        =   519
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   517
      Left            =   1200
      TabIndex        =   518
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   516
      Left            =   960
      TabIndex        =   517
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   515
      Left            =   720
      TabIndex        =   516
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   514
      Left            =   480
      TabIndex        =   515
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   513
      Left            =   240
      TabIndex        =   514
      Top             =   4800
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   512
      Left            =   0
      TabIndex        =   513
      Top             =   4800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   511
      Left            =   7440
      TabIndex        =   512
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   510
      Left            =   7200
      TabIndex        =   511
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   509
      Left            =   6960
      TabIndex        =   510
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   508
      Left            =   6720
      TabIndex        =   509
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   507
      Left            =   6480
      TabIndex        =   508
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   506
      Left            =   6240
      TabIndex        =   507
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   505
      Left            =   6000
      TabIndex        =   506
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   504
      Left            =   5760
      TabIndex        =   505
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   503
      Left            =   5520
      TabIndex        =   504
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   502
      Left            =   5280
      TabIndex        =   503
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   501
      Left            =   5040
      TabIndex        =   502
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   500
      Left            =   4800
      TabIndex        =   501
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   499
      Left            =   4560
      TabIndex        =   500
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   498
      Left            =   4320
      TabIndex        =   499
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   497
      Left            =   4080
      TabIndex        =   498
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   496
      Left            =   3840
      TabIndex        =   497
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   495
      Left            =   3600
      TabIndex        =   496
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   494
      Left            =   3360
      TabIndex        =   495
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   493
      Left            =   3120
      TabIndex        =   494
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   492
      Left            =   2880
      TabIndex        =   493
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   491
      Left            =   2640
      TabIndex        =   492
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   490
      Left            =   2400
      TabIndex        =   491
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   489
      Left            =   2160
      TabIndex        =   490
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   488
      Left            =   1920
      TabIndex        =   489
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   487
      Left            =   1680
      TabIndex        =   488
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   486
      Left            =   1440
      TabIndex        =   487
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   485
      Left            =   1200
      TabIndex        =   486
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   484
      Left            =   960
      TabIndex        =   485
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   483
      Left            =   720
      TabIndex        =   484
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   482
      Left            =   480
      TabIndex        =   483
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   481
      Left            =   240
      TabIndex        =   482
      Top             =   4560
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   480
      Left            =   0
      TabIndex        =   481
      Top             =   4560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   479
      Left            =   7440
      TabIndex        =   480
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   478
      Left            =   7200
      TabIndex        =   479
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   477
      Left            =   6960
      TabIndex        =   478
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   476
      Left            =   6720
      TabIndex        =   477
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   475
      Left            =   6480
      TabIndex        =   476
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   474
      Left            =   6240
      TabIndex        =   475
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   473
      Left            =   6000
      TabIndex        =   474
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   472
      Left            =   5760
      TabIndex        =   473
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   471
      Left            =   5520
      TabIndex        =   472
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   470
      Left            =   5280
      TabIndex        =   471
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   469
      Left            =   5040
      TabIndex        =   470
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   468
      Left            =   4800
      TabIndex        =   469
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   467
      Left            =   4560
      TabIndex        =   468
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   466
      Left            =   4320
      TabIndex        =   467
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   465
      Left            =   4080
      TabIndex        =   466
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   464
      Left            =   3840
      TabIndex        =   465
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   463
      Left            =   3600
      TabIndex        =   464
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   462
      Left            =   3360
      TabIndex        =   463
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   461
      Left            =   3120
      TabIndex        =   462
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   460
      Left            =   2880
      TabIndex        =   461
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   459
      Left            =   2640
      TabIndex        =   460
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   458
      Left            =   2400
      TabIndex        =   459
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   457
      Left            =   2160
      TabIndex        =   458
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   456
      Left            =   1920
      TabIndex        =   457
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   455
      Left            =   1680
      TabIndex        =   456
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   454
      Left            =   1440
      TabIndex        =   455
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   453
      Left            =   1200
      TabIndex        =   454
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   452
      Left            =   960
      TabIndex        =   453
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   451
      Left            =   720
      TabIndex        =   452
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   450
      Left            =   480
      TabIndex        =   451
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   449
      Left            =   240
      TabIndex        =   450
      Top             =   4320
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   448
      Left            =   0
      TabIndex        =   449
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   447
      Left            =   7440
      TabIndex        =   448
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   446
      Left            =   7200
      TabIndex        =   447
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   445
      Left            =   6960
      TabIndex        =   446
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   444
      Left            =   6720
      TabIndex        =   445
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   443
      Left            =   6480
      TabIndex        =   444
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   442
      Left            =   6240
      TabIndex        =   443
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   441
      Left            =   6000
      TabIndex        =   442
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   440
      Left            =   5760
      TabIndex        =   441
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   439
      Left            =   5520
      TabIndex        =   440
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   438
      Left            =   5280
      TabIndex        =   439
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   437
      Left            =   5040
      TabIndex        =   438
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   436
      Left            =   4800
      TabIndex        =   437
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   435
      Left            =   4560
      TabIndex        =   436
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   434
      Left            =   4320
      TabIndex        =   435
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   433
      Left            =   4080
      TabIndex        =   434
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   432
      Left            =   3840
      TabIndex        =   433
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   431
      Left            =   3600
      TabIndex        =   432
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   430
      Left            =   3360
      TabIndex        =   431
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   429
      Left            =   3120
      TabIndex        =   430
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   428
      Left            =   2880
      TabIndex        =   429
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   427
      Left            =   2640
      TabIndex        =   428
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   426
      Left            =   2400
      TabIndex        =   427
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   425
      Left            =   2160
      TabIndex        =   426
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   424
      Left            =   1920
      TabIndex        =   425
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   423
      Left            =   1680
      TabIndex        =   424
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   422
      Left            =   1440
      TabIndex        =   423
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   421
      Left            =   1200
      TabIndex        =   422
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   420
      Left            =   960
      TabIndex        =   421
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   419
      Left            =   720
      TabIndex        =   420
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   418
      Left            =   480
      TabIndex        =   419
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   417
      Left            =   240
      TabIndex        =   418
      Top             =   4080
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   416
      Left            =   0
      TabIndex        =   417
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   415
      Left            =   7440
      TabIndex        =   416
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   414
      Left            =   7200
      TabIndex        =   415
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   413
      Left            =   6960
      TabIndex        =   414
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   412
      Left            =   6720
      TabIndex        =   413
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   411
      Left            =   6480
      TabIndex        =   412
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   410
      Left            =   6240
      TabIndex        =   411
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   409
      Left            =   6000
      TabIndex        =   410
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   408
      Left            =   5760
      TabIndex        =   409
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   407
      Left            =   5520
      TabIndex        =   408
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   406
      Left            =   5280
      TabIndex        =   407
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   405
      Left            =   5040
      TabIndex        =   406
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   404
      Left            =   4800
      TabIndex        =   405
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   403
      Left            =   4560
      TabIndex        =   404
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   402
      Left            =   4320
      TabIndex        =   403
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   401
      Left            =   4080
      TabIndex        =   402
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   400
      Left            =   3840
      TabIndex        =   401
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   399
      Left            =   3600
      TabIndex        =   400
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   398
      Left            =   3360
      TabIndex        =   399
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   397
      Left            =   3120
      TabIndex        =   398
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   396
      Left            =   2880
      TabIndex        =   397
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   395
      Left            =   2640
      TabIndex        =   396
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   394
      Left            =   2400
      TabIndex        =   395
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   393
      Left            =   2160
      TabIndex        =   394
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   392
      Left            =   1920
      TabIndex        =   393
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   391
      Left            =   1680
      TabIndex        =   392
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   390
      Left            =   1440
      TabIndex        =   391
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   389
      Left            =   1200
      TabIndex        =   390
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   388
      Left            =   960
      TabIndex        =   389
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   387
      Left            =   720
      TabIndex        =   388
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   386
      Left            =   480
      TabIndex        =   387
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   385
      Left            =   240
      TabIndex        =   386
      Top             =   3840
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   384
      Left            =   0
      TabIndex        =   385
      Top             =   3840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   383
      Left            =   7440
      TabIndex        =   384
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   382
      Left            =   7200
      TabIndex        =   383
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   381
      Left            =   6960
      TabIndex        =   382
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   380
      Left            =   6720
      TabIndex        =   381
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   379
      Left            =   6480
      TabIndex        =   380
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   378
      Left            =   6240
      TabIndex        =   379
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   377
      Left            =   6000
      TabIndex        =   378
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   376
      Left            =   5760
      TabIndex        =   377
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   375
      Left            =   5520
      TabIndex        =   376
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   374
      Left            =   5280
      TabIndex        =   375
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   373
      Left            =   5040
      TabIndex        =   374
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   372
      Left            =   4800
      TabIndex        =   373
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   371
      Left            =   4560
      TabIndex        =   372
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   370
      Left            =   4320
      TabIndex        =   371
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   369
      Left            =   4080
      TabIndex        =   370
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   368
      Left            =   3840
      TabIndex        =   369
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   367
      Left            =   3600
      TabIndex        =   368
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   366
      Left            =   3360
      TabIndex        =   367
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   365
      Left            =   3120
      TabIndex        =   366
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   364
      Left            =   2880
      TabIndex        =   365
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   363
      Left            =   2640
      TabIndex        =   364
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   362
      Left            =   2400
      TabIndex        =   363
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   361
      Left            =   2160
      TabIndex        =   362
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   360
      Left            =   1920
      TabIndex        =   361
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   359
      Left            =   1680
      TabIndex        =   360
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   358
      Left            =   1440
      TabIndex        =   359
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   357
      Left            =   1200
      TabIndex        =   358
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   356
      Left            =   960
      TabIndex        =   357
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   355
      Left            =   720
      TabIndex        =   356
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   354
      Left            =   480
      TabIndex        =   355
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   353
      Left            =   240
      TabIndex        =   354
      Top             =   3600
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   352
      Left            =   0
      TabIndex        =   353
      Top             =   3600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   351
      Left            =   7440
      TabIndex        =   352
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   350
      Left            =   7200
      TabIndex        =   351
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   349
      Left            =   6960
      TabIndex        =   350
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   348
      Left            =   6720
      TabIndex        =   349
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   347
      Left            =   6480
      TabIndex        =   348
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   346
      Left            =   6240
      TabIndex        =   347
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   345
      Left            =   6000
      TabIndex        =   346
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   344
      Left            =   5760
      TabIndex        =   345
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   343
      Left            =   5520
      TabIndex        =   344
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   342
      Left            =   5280
      TabIndex        =   343
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   341
      Left            =   5040
      TabIndex        =   342
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   340
      Left            =   4800
      TabIndex        =   341
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   339
      Left            =   4560
      TabIndex        =   340
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   338
      Left            =   4320
      TabIndex        =   339
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   337
      Left            =   4080
      TabIndex        =   338
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   336
      Left            =   3840
      TabIndex        =   337
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   335
      Left            =   3600
      TabIndex        =   336
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   334
      Left            =   3360
      TabIndex        =   335
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   333
      Left            =   3120
      TabIndex        =   334
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   332
      Left            =   2880
      TabIndex        =   333
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   331
      Left            =   2640
      TabIndex        =   332
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   330
      Left            =   2400
      TabIndex        =   331
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   329
      Left            =   2160
      TabIndex        =   330
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   328
      Left            =   1920
      TabIndex        =   329
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   327
      Left            =   1680
      TabIndex        =   328
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   326
      Left            =   1440
      TabIndex        =   327
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   325
      Left            =   1200
      TabIndex        =   326
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   324
      Left            =   960
      TabIndex        =   325
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   323
      Left            =   720
      TabIndex        =   324
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   322
      Left            =   480
      TabIndex        =   323
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   321
      Left            =   240
      TabIndex        =   322
      Top             =   3360
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   320
      Left            =   0
      TabIndex        =   321
      Top             =   3360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   319
      Left            =   7440
      TabIndex        =   320
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   318
      Left            =   7200
      TabIndex        =   319
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   317
      Left            =   6960
      TabIndex        =   318
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   316
      Left            =   6720
      TabIndex        =   317
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   315
      Left            =   6480
      TabIndex        =   316
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   314
      Left            =   6240
      TabIndex        =   315
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   313
      Left            =   6000
      TabIndex        =   314
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   312
      Left            =   5760
      TabIndex        =   313
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   311
      Left            =   5520
      TabIndex        =   312
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   310
      Left            =   5280
      TabIndex        =   311
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   309
      Left            =   5040
      TabIndex        =   310
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   308
      Left            =   4800
      TabIndex        =   309
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   307
      Left            =   4560
      TabIndex        =   308
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   306
      Left            =   4320
      TabIndex        =   307
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   305
      Left            =   4080
      TabIndex        =   306
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   304
      Left            =   3840
      TabIndex        =   305
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   303
      Left            =   3600
      TabIndex        =   304
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   302
      Left            =   3360
      TabIndex        =   303
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   301
      Left            =   3120
      TabIndex        =   302
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   300
      Left            =   2880
      TabIndex        =   301
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   299
      Left            =   2640
      TabIndex        =   300
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   298
      Left            =   2400
      TabIndex        =   299
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   297
      Left            =   2160
      TabIndex        =   298
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   296
      Left            =   1920
      TabIndex        =   297
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   295
      Left            =   1680
      TabIndex        =   296
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   294
      Left            =   1440
      TabIndex        =   295
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   293
      Left            =   1200
      TabIndex        =   294
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   292
      Left            =   960
      TabIndex        =   293
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   291
      Left            =   720
      TabIndex        =   292
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   290
      Left            =   480
      TabIndex        =   291
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   289
      Left            =   240
      TabIndex        =   290
      Top             =   3120
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   288
      Left            =   0
      TabIndex        =   289
      Top             =   3120
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   287
      Left            =   7440
      TabIndex        =   288
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   286
      Left            =   7200
      TabIndex        =   287
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   285
      Left            =   6960
      TabIndex        =   286
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   284
      Left            =   6720
      TabIndex        =   285
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   283
      Left            =   6480
      TabIndex        =   284
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   282
      Left            =   6240
      TabIndex        =   283
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   281
      Left            =   6000
      TabIndex        =   282
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   280
      Left            =   5760
      TabIndex        =   281
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   279
      Left            =   5520
      TabIndex        =   280
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   278
      Left            =   5280
      TabIndex        =   279
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   277
      Left            =   5040
      TabIndex        =   278
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   276
      Left            =   4800
      TabIndex        =   277
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   275
      Left            =   4560
      TabIndex        =   276
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   274
      Left            =   4320
      TabIndex        =   275
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   273
      Left            =   4080
      TabIndex        =   274
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   272
      Left            =   3840
      TabIndex        =   273
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   271
      Left            =   3600
      TabIndex        =   272
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   270
      Left            =   3360
      TabIndex        =   271
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   269
      Left            =   3120
      TabIndex        =   270
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   268
      Left            =   2880
      TabIndex        =   269
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   267
      Left            =   2640
      TabIndex        =   268
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   266
      Left            =   2400
      TabIndex        =   267
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   265
      Left            =   2160
      TabIndex        =   266
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   264
      Left            =   1920
      TabIndex        =   265
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   263
      Left            =   1680
      TabIndex        =   264
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   262
      Left            =   1440
      TabIndex        =   263
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   261
      Left            =   1200
      TabIndex        =   262
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   260
      Left            =   960
      TabIndex        =   261
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   259
      Left            =   720
      TabIndex        =   260
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   258
      Left            =   480
      TabIndex        =   259
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   257
      Left            =   240
      TabIndex        =   258
      Top             =   2880
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   256
      Left            =   0
      TabIndex        =   257
      Top             =   2880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   255
      Left            =   7440
      TabIndex        =   256
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   254
      Left            =   7200
      TabIndex        =   255
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   253
      Left            =   6960
      TabIndex        =   254
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   252
      Left            =   6720
      TabIndex        =   253
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   251
      Left            =   6480
      TabIndex        =   252
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   250
      Left            =   6240
      TabIndex        =   251
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   249
      Left            =   6000
      TabIndex        =   250
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   248
      Left            =   5760
      TabIndex        =   249
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   247
      Left            =   5520
      TabIndex        =   248
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   246
      Left            =   5280
      TabIndex        =   247
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   245
      Left            =   5040
      TabIndex        =   246
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   244
      Left            =   4800
      TabIndex        =   245
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   243
      Left            =   4560
      TabIndex        =   244
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   242
      Left            =   4320
      TabIndex        =   243
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   241
      Left            =   4080
      TabIndex        =   242
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   240
      Left            =   3840
      TabIndex        =   241
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   239
      Left            =   3600
      TabIndex        =   240
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   238
      Left            =   3360
      TabIndex        =   239
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   237
      Left            =   3120
      TabIndex        =   238
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   236
      Left            =   2880
      TabIndex        =   237
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   235
      Left            =   2640
      TabIndex        =   236
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   234
      Left            =   2400
      TabIndex        =   235
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   233
      Left            =   2160
      TabIndex        =   234
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   232
      Left            =   1920
      TabIndex        =   233
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   231
      Left            =   1680
      TabIndex        =   232
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   230
      Left            =   1440
      TabIndex        =   231
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   229
      Left            =   1200
      TabIndex        =   230
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   228
      Left            =   960
      TabIndex        =   229
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   227
      Left            =   720
      TabIndex        =   228
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   226
      Left            =   480
      TabIndex        =   227
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   225
      Left            =   240
      TabIndex        =   226
      Top             =   2640
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   224
      Left            =   0
      TabIndex        =   225
      Top             =   2640
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   223
      Left            =   7440
      TabIndex        =   224
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   222
      Left            =   7200
      TabIndex        =   223
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   221
      Left            =   6960
      TabIndex        =   222
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   220
      Left            =   6720
      TabIndex        =   221
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   219
      Left            =   6480
      TabIndex        =   220
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   218
      Left            =   6240
      TabIndex        =   219
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   217
      Left            =   6000
      TabIndex        =   218
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   216
      Left            =   5760
      TabIndex        =   217
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   215
      Left            =   5520
      TabIndex        =   216
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   214
      Left            =   5280
      TabIndex        =   215
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   213
      Left            =   5040
      TabIndex        =   214
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   212
      Left            =   4800
      TabIndex        =   213
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   211
      Left            =   4560
      TabIndex        =   212
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   210
      Left            =   4320
      TabIndex        =   211
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   209
      Left            =   4080
      TabIndex        =   210
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   208
      Left            =   3840
      TabIndex        =   209
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   207
      Left            =   3600
      TabIndex        =   208
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   206
      Left            =   3360
      TabIndex        =   207
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   205
      Left            =   3120
      TabIndex        =   206
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   204
      Left            =   2880
      TabIndex        =   205
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   203
      Left            =   2640
      TabIndex        =   204
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   202
      Left            =   2400
      TabIndex        =   203
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   201
      Left            =   2160
      TabIndex        =   202
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   200
      Left            =   1920
      TabIndex        =   201
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   199
      Left            =   1680
      TabIndex        =   200
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   198
      Left            =   1440
      TabIndex        =   199
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   197
      Left            =   1200
      TabIndex        =   198
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   196
      Left            =   960
      TabIndex        =   197
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   195
      Left            =   720
      TabIndex        =   196
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   194
      Left            =   480
      TabIndex        =   195
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   193
      Left            =   240
      TabIndex        =   194
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   192
      Left            =   0
      TabIndex        =   193
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   191
      Left            =   7440
      TabIndex        =   192
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   190
      Left            =   7200
      TabIndex        =   191
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   189
      Left            =   6960
      TabIndex        =   190
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   188
      Left            =   6720
      TabIndex        =   189
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   187
      Left            =   6480
      TabIndex        =   188
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   186
      Left            =   6240
      TabIndex        =   187
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   185
      Left            =   6000
      TabIndex        =   186
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   184
      Left            =   5760
      TabIndex        =   185
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   183
      Left            =   5520
      TabIndex        =   184
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   182
      Left            =   5280
      TabIndex        =   183
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   181
      Left            =   5040
      TabIndex        =   182
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   180
      Left            =   4800
      TabIndex        =   181
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   179
      Left            =   4560
      TabIndex        =   180
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   178
      Left            =   4320
      TabIndex        =   179
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   177
      Left            =   4080
      TabIndex        =   178
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   176
      Left            =   3840
      TabIndex        =   177
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   175
      Left            =   3600
      TabIndex        =   176
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   174
      Left            =   3360
      TabIndex        =   175
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   173
      Left            =   3120
      TabIndex        =   174
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   172
      Left            =   2880
      TabIndex        =   173
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   171
      Left            =   2640
      TabIndex        =   172
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   170
      Left            =   2400
      TabIndex        =   171
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   169
      Left            =   2160
      TabIndex        =   170
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   168
      Left            =   1920
      TabIndex        =   169
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   167
      Left            =   1680
      TabIndex        =   168
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   166
      Left            =   1440
      TabIndex        =   167
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   165
      Left            =   1200
      TabIndex        =   166
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   164
      Left            =   960
      TabIndex        =   165
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   163
      Left            =   720
      TabIndex        =   164
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   162
      Left            =   480
      TabIndex        =   163
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   161
      Left            =   240
      TabIndex        =   162
      Top             =   2160
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   160
      Left            =   0
      TabIndex        =   161
      Top             =   2160
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   159
      Left            =   7440
      TabIndex        =   160
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   158
      Left            =   7200
      TabIndex        =   159
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   157
      Left            =   6960
      TabIndex        =   158
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   156
      Left            =   6720
      TabIndex        =   157
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   155
      Left            =   6480
      TabIndex        =   156
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   154
      Left            =   6240
      TabIndex        =   155
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   153
      Left            =   6000
      TabIndex        =   154
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   152
      Left            =   5760
      TabIndex        =   153
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   151
      Left            =   5520
      TabIndex        =   152
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   150
      Left            =   5280
      TabIndex        =   151
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   149
      Left            =   5040
      TabIndex        =   150
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   148
      Left            =   4800
      TabIndex        =   149
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   147
      Left            =   4560
      TabIndex        =   148
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   146
      Left            =   4320
      TabIndex        =   147
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   145
      Left            =   4080
      TabIndex        =   146
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   144
      Left            =   3840
      TabIndex        =   145
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   143
      Left            =   3600
      TabIndex        =   144
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   142
      Left            =   3360
      TabIndex        =   143
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   141
      Left            =   3120
      TabIndex        =   142
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   140
      Left            =   2880
      TabIndex        =   141
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   139
      Left            =   2640
      TabIndex        =   140
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   138
      Left            =   2400
      TabIndex        =   139
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   137
      Left            =   2160
      TabIndex        =   138
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   136
      Left            =   1920
      TabIndex        =   137
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   135
      Left            =   1680
      TabIndex        =   136
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   134
      Left            =   1440
      TabIndex        =   135
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   133
      Left            =   1200
      TabIndex        =   134
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   132
      Left            =   960
      TabIndex        =   133
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   131
      Left            =   720
      TabIndex        =   132
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   130
      Left            =   480
      TabIndex        =   131
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   129
      Left            =   240
      TabIndex        =   130
      Top             =   1920
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   128
      Left            =   0
      TabIndex        =   129
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   127
      Left            =   7440
      TabIndex        =   128
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   126
      Left            =   7200
      TabIndex        =   127
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   125
      Left            =   6960
      TabIndex        =   126
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   124
      Left            =   6720
      TabIndex        =   125
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   123
      Left            =   6480
      TabIndex        =   124
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   122
      Left            =   6240
      TabIndex        =   123
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   121
      Left            =   6000
      TabIndex        =   122
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   120
      Left            =   5760
      TabIndex        =   121
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   119
      Left            =   5520
      TabIndex        =   120
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   118
      Left            =   5280
      TabIndex        =   119
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   117
      Left            =   5040
      TabIndex        =   118
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   116
      Left            =   4800
      TabIndex        =   117
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   115
      Left            =   4560
      TabIndex        =   116
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   114
      Left            =   4320
      TabIndex        =   115
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   113
      Left            =   4080
      TabIndex        =   114
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   112
      Left            =   3840
      TabIndex        =   113
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   111
      Left            =   3600
      TabIndex        =   112
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   110
      Left            =   3360
      TabIndex        =   111
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   109
      Left            =   3120
      TabIndex        =   110
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   108
      Left            =   2880
      TabIndex        =   109
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   107
      Left            =   2640
      TabIndex        =   108
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   106
      Left            =   2400
      TabIndex        =   107
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   105
      Left            =   2160
      TabIndex        =   106
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   104
      Left            =   1920
      TabIndex        =   105
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   103
      Left            =   1680
      TabIndex        =   104
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   102
      Left            =   1440
      TabIndex        =   103
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   101
      Left            =   1200
      TabIndex        =   102
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   100
      Left            =   960
      TabIndex        =   101
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   99
      Left            =   720
      TabIndex        =   100
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   98
      Left            =   480
      TabIndex        =   99
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   97
      Left            =   240
      TabIndex        =   98
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   96
      Left            =   0
      TabIndex        =   97
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   95
      Left            =   7440
      TabIndex        =   96
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   94
      Left            =   7200
      TabIndex        =   95
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   93
      Left            =   6960
      TabIndex        =   94
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   92
      Left            =   6720
      TabIndex        =   93
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   91
      Left            =   6480
      TabIndex        =   92
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   90
      Left            =   6240
      TabIndex        =   91
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   89
      Left            =   6000
      TabIndex        =   90
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   88
      Left            =   5760
      TabIndex        =   89
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   87
      Left            =   5520
      TabIndex        =   88
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   86
      Left            =   5280
      TabIndex        =   87
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   85
      Left            =   5040
      TabIndex        =   86
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   84
      Left            =   4800
      TabIndex        =   85
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   83
      Left            =   4560
      TabIndex        =   84
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   82
      Left            =   4320
      TabIndex        =   83
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   81
      Left            =   4080
      TabIndex        =   82
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   80
      Left            =   3840
      TabIndex        =   81
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   79
      Left            =   3600
      TabIndex        =   80
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   78
      Left            =   3360
      TabIndex        =   79
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   77
      Left            =   3120
      TabIndex        =   78
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   76
      Left            =   2880
      TabIndex        =   77
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   75
      Left            =   2640
      TabIndex        =   76
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   74
      Left            =   2400
      TabIndex        =   75
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   73
      Left            =   2160
      TabIndex        =   74
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   72
      Left            =   1920
      TabIndex        =   73
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   71
      Left            =   1680
      TabIndex        =   72
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   70
      Left            =   1440
      TabIndex        =   71
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   69
      Left            =   1200
      TabIndex        =   70
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   68
      Left            =   960
      TabIndex        =   69
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   67
      Left            =   720
      TabIndex        =   68
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   66
      Left            =   480
      TabIndex        =   67
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   65
      Left            =   240
      TabIndex        =   66
      Top             =   1440
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   64
      Left            =   0
      TabIndex        =   65
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   63
      Left            =   7440
      TabIndex        =   64
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   62
      Left            =   7200
      TabIndex        =   63
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   61
      Left            =   6960
      TabIndex        =   62
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   60
      Left            =   6720
      TabIndex        =   61
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   59
      Left            =   6480
      TabIndex        =   60
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   58
      Left            =   6240
      TabIndex        =   59
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   57
      Left            =   6000
      TabIndex        =   58
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   56
      Left            =   5760
      TabIndex        =   57
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   55
      Left            =   5520
      TabIndex        =   56
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   54
      Left            =   5280
      TabIndex        =   55
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   53
      Left            =   5040
      TabIndex        =   54
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   52
      Left            =   4800
      TabIndex        =   53
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   51
      Left            =   4560
      TabIndex        =   52
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   50
      Left            =   4320
      TabIndex        =   51
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   49
      Left            =   4080
      TabIndex        =   50
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   48
      Left            =   3840
      TabIndex        =   49
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   47
      Left            =   3600
      TabIndex        =   48
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   46
      Left            =   3360
      TabIndex        =   47
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   45
      Left            =   3120
      TabIndex        =   46
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   44
      Left            =   2880
      TabIndex        =   45
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   43
      Left            =   2640
      TabIndex        =   44
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   42
      Left            =   2400
      TabIndex        =   43
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   41
      Left            =   2160
      TabIndex        =   42
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   40
      Left            =   1920
      TabIndex        =   41
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   39
      Left            =   1680
      TabIndex        =   40
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   38
      Left            =   1440
      TabIndex        =   39
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   37
      Left            =   1200
      TabIndex        =   38
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   36
      Left            =   960
      TabIndex        =   37
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   35
      Left            =   720
      TabIndex        =   36
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   34
      Left            =   480
      TabIndex        =   35
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   33
      Left            =   240
      TabIndex        =   34
      Top             =   1200
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   32
      Left            =   0
      TabIndex        =   33
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   31
      Left            =   7440
      TabIndex        =   32
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   30
      Left            =   7200
      TabIndex        =   31
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   29
      Left            =   6960
      TabIndex        =   30
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   28
      Left            =   6720
      TabIndex        =   29
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   27
      Left            =   6480
      TabIndex        =   28
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   26
      Left            =   6240
      TabIndex        =   27
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   25
      Left            =   6000
      TabIndex        =   26
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   24
      Left            =   5760
      TabIndex        =   25
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   23
      Left            =   5520
      TabIndex        =   24
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   22
      Left            =   5280
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   21
      Left            =   5040
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   20
      Left            =   4800
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   19
      Left            =   4560
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   18
      Left            =   4320
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   17
      Left            =   4080
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   16
      Left            =   3840
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   15
      Left            =   3600
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   14
      Left            =   3360
      TabIndex        =   15
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   13
      Left            =   3120
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   12
      Left            =   2880
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   11
      Left            =   2640
      TabIndex        =   12
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   8
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   7
      Left            =   1680
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   6
      Left            =   1440
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton Tile 
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Menu mGame 
      Caption         =   "游戏(&G)"
      Begin VB.Menu mStart 
         Caption         =   "开始(&S)"
      End
      Begin VB.Menu mExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnStart_Click()
    Call NewGame
    'start = True
End Sub

Private Sub mAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mExit_Click()
    End
End Sub

Private Sub mStart_Click()
    Call NewGame
End Sub

Private Sub Tile_Click(Index As Integer)
    opennum = 0
    start = True
    Call opentile(Index)
    For i = 0 To 639
        If Tile(i).Caption <> "" And Tile(i).Caption <> "●" Then
            opennum = opennum + 1
        End If
    Next
    If opennum + num = 540 Then
        frmWin.Show 1
        Call NewGame
    End If
End Sub
