package excel

import (
	"fmt"
	"os"
	"strconv"

	"github.com/360EntSecGroup-Skylar/excelize"
)

//读取 平台补贴整合模板 (1).xlsx
func Readbook5(excel map[int]string) ([]string, []string) {

	book5, err1 := excelize.OpenFile("./" + excel[6])
	if err1 != nil {
		fmt.Println(err1)
	}
	// rows, _ := book5.GetRows(book5.GetSheetName(1))

	// fmt.Println(rows[2][1])
	// fmt.Println(rows[1][0])
	arr1 := make([]string, 0, 50)
	arr2 := make([]string, 0, 50)
	for i := 3; i < 46; i++ {
		cellB, _ := book5.GetCellValue(book5.GetSheetName(1), ("B" + strconv.Itoa(i)))
		cellC, _ := book5.GetCellValue(book5.GetSheetName(1), ("C" + strconv.Itoa(i)))
		arr1 = append(arr1, cellB)
		arr2 = append(arr2, cellC)
		// arr1 = append(arr1, rows[i][1])
		// arr2 = append(arr2, rows[i][2])
	}
	// fmt.Println(arr1[0])
	// fmt.Println(arr2[1])
	return arr1, arr2

}

//读取 分平台营业数据 (解绑后模板-新) (1).xlsx
func Readbook1(excel map[int]string) map[int][]string {

	var arrs map[int][]string
	arrs = make(map[int][]string)

	defer func() {
		err := recover()
		if err != nil {
			//说明捕获到异常
			fmt.Print("err=", err)
			fmt.Println(len(arrs))
		}
	}()

	book1, err1 := excelize.OpenFile("./" + excel[2])
	if err1 != nil {
		fmt.Println(err1)
	}
	// rows, _ := book1.GetRows(book1.GetSheetName(1))

	fmt.Println(book1.GetCellValue(book1.GetSheetName(1), "B42"))
	for i := 2; i < 44; i++ {
		arr1 := make([]string, 0, 16)
		// f64, err := strconv.ParseFloat(rows[i][2], 64)
		// if err != nil {
		// 	fmt.Println(err)
		// }
		// fmt.Println(f64)
		// fmt.Printf("v1 type:%T\n", f64)
		// fmt.Println(rows[i][2])
		cellC, _ := book1.GetCellValue(book1.GetSheetName(1), ("C" + strconv.Itoa(i)))
		cellD, _ := book1.GetCellValue(book1.GetSheetName(1), ("D" + strconv.Itoa(i)))
		cellE, _ := book1.GetCellValue(book1.GetSheetName(1), ("E" + strconv.Itoa(i)))
		cellK, _ := book1.GetCellValue(book1.GetSheetName(1), ("K" + strconv.Itoa(i)))
		cellL, _ := book1.GetCellValue(book1.GetSheetName(1), ("L" + strconv.Itoa(i)))
		cellO, _ := book1.GetCellValue(book1.GetSheetName(1), ("O" + strconv.Itoa(i)))
		cellQ, _ := book1.GetCellValue(book1.GetSheetName(1), ("Q" + strconv.Itoa(i)))
		cellT, _ := book1.GetCellValue(book1.GetSheetName(1), ("T" + strconv.Itoa(i)))
		cellU, _ := book1.GetCellValue(book1.GetSheetName(1), ("U" + strconv.Itoa(i)))
		cellV, _ := book1.GetCellValue(book1.GetSheetName(1), ("V" + strconv.Itoa(i)))
		cellAB, _ := book1.GetCellValue(book1.GetSheetName(1), ("AB" + strconv.Itoa(i)))
		cellAC, _ := book1.GetCellValue(book1.GetSheetName(1), ("AC" + strconv.Itoa(i)))
		cellAD, _ := book1.GetCellValue(book1.GetSheetName(1), ("AD" + strconv.Itoa(i)))
		cellAE, _ := book1.GetCellValue(book1.GetSheetName(1), ("AE" + strconv.Itoa(i)))
		cellAH, _ := book1.GetCellValue(book1.GetSheetName(1), ("AH" + strconv.Itoa(i)))
		cellAJ, _ := book1.GetCellValue(book1.GetSheetName(1), ("AJ" + strconv.Itoa(i)))

		arr1 = append(arr1, cellC)
		arr1 = append(arr1, cellD)
		arr1 = append(arr1, cellE)
		arr1 = append(arr1, cellK)
		arr1 = append(arr1, cellL)
		arr1 = append(arr1, cellO)
		arr1 = append(arr1, cellQ)
		arr1 = append(arr1, cellT)
		arr1 = append(arr1, cellU)
		arr1 = append(arr1, cellV)
		arr1 = append(arr1, cellAB)
		arr1 = append(arr1, cellAC)
		arr1 = append(arr1, cellAD)
		arr1 = append(arr1, cellAE)
		arr1 = append(arr1, cellAH)
		arr1 = append(arr1, cellAJ)

		// arr1 = append(arr1, rows[i][2])
		// arr1 = append(arr1, rows[i][3])
		// arr1 = append(arr1, rows[i][4])
		// arr1 = append(arr1, rows[i][10])
		// arr1 = append(arr1, rows[i][11])
		// arr1 = append(arr1, rows[i][14])
		// arr1 = append(arr1, rows[i][16])
		// arr1 = append(arr1, rows[i][19])
		// arr1 = append(arr1, rows[i][20])
		// arr1 = append(arr1, rows[i][21])
		// arr1 = append(arr1, rows[i][27])
		// arr1 = append(arr1, rows[i][28])
		// arr1 = append(arr1, rows[i][29])
		// arr1 = append(arr1, rows[i][30])
		// arr1 = append(arr1, rows[i][33])
		// arr1 = append(arr1, rows[i][35])
		// fmt.Println(arr1)
		arrs[i] = arr1
	}

	// fmt.Println(arrs[37])
	return arrs
}

//读取 评论率 (68).xlsx
func Readbook2(excel map[int]string) map[int][]string {
	var arrs map[int][]string
	arrs = make(map[int][]string)

	defer func() {
		err := recover()
		if err != nil {
			//说明捕获到异常
			fmt.Print("err=", err)
			fmt.Println(len(arrs))
			if err == "runtime error: index out of range" {
				fmt.Println(arrs[37])
			}
		}
	}()

	book2, err1 := excelize.OpenFile("./" + excel[3])
	if err1 != nil {
		fmt.Println(err1)
	}
	// rows, _ := book2.GetRows(book2.GetSheetName(1))

	for i := 2; i < 44; i++ {
		arr1 := make([]string, 0, 3)
		// fmt.Println(f64)
		// fmt.Printf("v1 type:%T\n", f64)
		cellC, _ := book2.GetCellValue(book2.GetSheetName(1), ("C" + strconv.Itoa(i)))
		cellD, _ := book2.GetCellValue(book2.GetSheetName(1), ("D" + strconv.Itoa(i)))
		cellE, _ := book2.GetCellValue(book2.GetSheetName(1), ("E" + strconv.Itoa(i)))
		arr1 = append(arr1, cellC)
		arr1 = append(arr1, cellD)
		arr1 = append(arr1, cellE)
		// fmt.Println(arr1)
		arrs[i] = arr1
	}

	return arrs

}

//保存 外送部数据记录表 10.02.xlsx
func Writebook3(excel map[int]string, arrs map[int][]string, arry1 []string, arry2 []string) {

	defer func() {
		err := recover()
		if err != nil {
			//说明捕获到异常
			fmt.Print("err=", err)
			fmt.Println(len(arrs))
			// if err == "runtime error: index out of range" {
			// 	fmt.Println(arrs[37])
			// }
		}
	}()

	book3, _ := excelize.OpenFile("./" + excel[4])
	for i := 2; i < 44; i++ {
		cellC, err1 := strconv.ParseFloat(arrs[i][0], 64)
		cellD, err2 := strconv.ParseFloat(arrs[i][1], 64)
		cellE, err3 := strconv.ParseFloat(arrs[i][2], 64)
		cellK, err4 := strconv.ParseFloat(arrs[i][3], 64)
		cellL, err5 := strconv.ParseFloat(arrs[i][4], 64)
		cellO, err6 := strconv.ParseFloat(arrs[i][5], 64)
		cellQ, err7 := strconv.ParseFloat(arrs[i][6], 64)
		cellT, err8 := strconv.ParseFloat(arrs[i][7], 64)
		cellU, err9 := strconv.ParseFloat(arrs[i][8], 64)
		cellV, err10 := strconv.ParseFloat(arrs[i][9], 64)
		cellAB, err11 := strconv.ParseFloat(arrs[i][10], 64)
		cellAC, err12 := strconv.ParseFloat(arrs[i][11], 64)
		cellAD, err13 := strconv.ParseFloat(arrs[i][12], 64)
		cellAE, err14 := strconv.ParseFloat(arrs[i][13], 64)
		cellAH, err15 := strconv.ParseFloat(arrs[i][14], 64)
		cellAJ, err16 := strconv.ParseFloat(arrs[i][15], 64)

		if err1 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "R"+excel[0], cellC)
		}
		if err2 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "S"+excel[0], cellD)
		}
		if err3 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "T"+excel[0], cellE)
		}
		if err4 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AA"+excel[0], cellK)
		}
		if err5 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AB"+excel[0], cellL)
		}
		if err6 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AE"+excel[0], cellO)
		}
		if err7 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AG"+excel[0], cellQ)
		}
		if err8 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AM"+excel[0], cellT)
		}
		if err9 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AN"+excel[0], cellU)
		}
		if err10 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AO"+excel[0], cellV)
		}
		if err11 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AV"+excel[0], cellAB)
		}
		if err12 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AW"+excel[0], cellAC)
		}
		if err13 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AX"+excel[0], cellAD)
		}
		if err14 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "AY"+excel[0], cellAE)
		}
		if err15 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "BB"+excel[0], cellAH)
		}
		if err16 == nil {
			book3.SetCellValue(book3.GetSheetName(i+2), "BD"+excel[0], cellAJ)
		}

		row, _ := strconv.Atoi(excel[0])
		row = row + 1

		book3.SetCellValue(book3.GetSheetName(i+2), "AJ"+strconv.Itoa(row), arry1[i-2])
		book3.SetCellValue(book3.GetSheetName(i+2), "BG"+strconv.Itoa(row), arry2[i-2])

	}
	book3.Save()
}

//保存外卖差评截图 10.02.xlsx
func Writebook4(excel map[int]string, arrs map[int][]string) {
	book4, _ := excelize.OpenFile("./" + excel[5])
	// number, _ := strconv.Atoi(excel[1])
	fmt.Println(arrs[2])
	fmt.Println("~~~~~~~~~~~~~~~")
	fmt.Println(len(book4.GetSheetMap()))
	if "" == book4.GetSheetName(100) {
		fmt.Println("一样一样的")
	}

	fmt.Println("~~~~~~~~~~~~~~~")

	for i := 2; i < 44; i++ {
		arr1, err1 := strconv.Atoi(arrs[i][0])
		arr2, err2 := strconv.Atoi(arrs[i][1])
		arr3, err3 := strconv.Atoi(arrs[i][2])
		if err1 == nil {
			book4.SetCellValue(book4.GetSheetName(i+1), "F"+excel[1], arr1)
		}
		if err2 == nil {
			book4.SetCellValue(book4.GetSheetName(i+1), "H"+excel[1], arr2)
		}
		if err3 == nil {
			book4.SetCellValue(book4.GetSheetName(i+1), "J"+excel[1], arr3)
		}

	}
	book4.Save()

}

func Main(excelname map[int]string) {
	// fileName := "./name.txt"
	// file, err := os.OpenFile(fileName, os.O_RDWR, 0666)
	// if err != nil {
	// 	fmt.Println("Open file error!", err)
	// 	return
	// }

	// buf := bufio.NewReader(file)
	// var excel map[int]string
	// excel = make(map[int]string)
	// i := 0
	// for {

	// 	line, err := buf.ReadString('\n')
	// 	line = strings.TrimSpace(line)
	// 	// fmt.Println(line)
	// 	excel[i] = line
	// 	i++
	// 	if err != nil {
	// 		if err == io.EOF {
	// 			fmt.Println("File read ok!")
	// 			break
	// 		} else {
	// 			fmt.Println("Read file error!", err)
	// 			return
	// 		}
	// 	}
	// }

	// fileName := "./name.txt"
	// file, err := os.OpenFile(fileName, os.O_RDWR, 0666)
	// if err != nil {
	// 	fmt.Println("Open file error!", err)
	// 	return
	// }

	// buf := bufio.NewReader(file)
	// var excelname map[int]string
	// excelname = make(map[int]string)
	// i := 0
	// for {

	// 	line, err := buf.ReadString('\n')
	// 	line = strings.TrimSpace(line)
	// 	// fmt.Println(line)
	// 	excelname[i] = line
	// 	i++
	// 	if err != nil {
	// 		if err == io.EOF {
	// 			fmt.Println("File read ok!")
	// 			break
	// 		} else {
	// 			fmt.Println("Read file error!", err)
	// 			return
	// 		}
	// 	}
	// }

	// fmt.Println(PathExists("./" + excelname[2]))
	fmt.Println(excelname)
	arry1, arry2 := Readbook5(excelname)
	arrs1 := Readbook1(excelname)
	arrs2 := Readbook2(excelname)
	Writebook4(excelname, arrs2)
	Writebook3(excelname, arrs1, arry1, arry2)

}

func PathExists(path string) (bool, error) {
	_, err := os.Stat(path)
	if err == nil {
		return true, nil
	}
	if os.IsNotExist(err) {
		return false, nil
	}
	return false, err
}
