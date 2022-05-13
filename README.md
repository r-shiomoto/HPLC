# HPLC

## Background
* I got bored with rearranging HPLC table based on rrt (stands for Relative Retention Time) and renaming file name which HPLC machine names as serial number.
* So, I developed scripts automating this boring tasks.

## Feature
* You can get excel file written rt, rrt, area and area% automatically.
* You have no need to rename file name to rearrange HPLC table
* You have option to only rename file name.

![outline](https://user-images.githubusercontent.com/57758623/167236803-dbbece7d-9b63-45f5-ab55-5bd3b71bc6ec.png)
## Python version
* python 3.7.11

## Requirement
* numpy version 1.20.3
* pandas version 1.3.4
* openpyxl version 3.0.9

## How to use
1. Run the script.
2. Enter path you want to rearrange.
3. Enter standard rt each file. The script propose it based on maximum area and paste it on clipboard. So you only paste and push enter if you want to rearrange profile of only product.
4. Get excel file excel file written rt, rrt, area and area%.

## Note
* There are scripts for waters and hitachi.
* Text of waters should be designed, referring to "sample.txt".
* Script of hitachi is not good because HPLC of hitachi is not major.

## Information
* Author : Ryuhei Shiomoto
* Created Date : 2022/05/07

## License
* 
