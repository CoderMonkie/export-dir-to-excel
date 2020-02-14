# Export Directory To Excel

tool for exporting directory information as excel

> Written in JavaScript

---

## ０、何だこれ？ About

本ツールは、以下の機能を提供している： 

- 指定されたフォルダパスを分析し、  
  フォルダ構造詳細をエクセルとして出力する

- ZIP ファイルに対しても中身構造を出力できる

※注：構造階層が深く、ファイルも多い場合を避けてください。  
　。。時間かかりすぎるから。。

---

## 1、前提 Precondition

1. Git

   if you dont want to install git,  
   you can download the tool directly.

   [Download-ExportDirToExcel](https://github.com/CoderMonkie/export-dir-to-excel/archive/master.zip)

   --return--

   download git from here:

   [git-downloads](https://git-scm.com/downloads)

   install by default.

   use git to get the source code of this tool.

   ```bat
   git clone https://github.com/CoderMonkie/export-dir-to-excel.git
   ```

2. Nodejs

   download from here:

   [Nodejs-HomePage](https://nodejs.org/en/)

   install by default.

---

## ２、使いかた Usage

0. Open the directory of this tool in cmd

   run commad below:

   ```bat
   npm install
   ```

   ※run it Only Once !!!

1. Doubly click [run.bat] to start  

2. [Must] Input path of target folder in cmd  
   support &lt;drag & drop>

3. [Optional] Input path for output or press Enter to use default   

    so easy?!

    ※Run 1~3 every time.

---

## ３、そのた Other

補足：

- 2.2 で press Enter、2.3 で &lt;path> 入力した場合、  
  以下となる：

  - Target Folder : &lt;path>
  - Output Path   : (default)