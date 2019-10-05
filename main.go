package main

import (
	"bufio"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"strings"

	"./excel"

	"github.com/lxn/walk"
	. "github.com/lxn/walk/declarative"
)

//go build -ldflags="-H windowsgui"
var isSpecialMode = walk.NewMutableCondition()

func main() {
	var inTE, outTE *walk.TextEdit
	var inbutthon *walk.PushButton
	// name := "1\r\n" +
	// 	"2\r\n" +
	// 	"3\r\n"
	fileName := "./name.txt"
	// file, err := os.OpenFile(fileName, os.O_RDWR, 0666)
	// if err != nil {
	// 	fmt.Println("Open file error!", err)
	// 	return
	// }
	buf, _ := ioutil.ReadFile(fileName)

	// buf := bufio.NewReader(file)
	fmt.Println(string(buf))
	MustRegisterCondition("isSpecialMode", isSpecialMode)
	isSpecialMode.SetSatisfied(true)

	MainWindow{
		Title:   "xiaochuan测试",
		MinSize: Size{600, 400},
		Layout:  VBox{},
		Children: []Widget{
			HSplitter{
				Children: []Widget{
					TextEdit{AssignTo: &inTE, Text: string(buf)},
					TextEdit{AssignTo: &outTE, ReadOnly: true, Text: "点完别动等着"},
				},
			},
			PushButton{
				Text:     "点这里",
				AssignTo: &inbutthon,
				Enabled:  Bind("isSpecialMode"),
				OnClicked: func() {
					// outTE.SetText("等着，别动")

					outTE.SetText(strings.ToUpper(inTE.Text()))
					r := strings.NewReader(inTE.Text())

					buf := bufio.NewReader(r)
					var excelname map[int]string
					excelname = make(map[int]string)
					i := 0
					for {

						line, err := buf.ReadString('\n')
						line = strings.TrimSpace(line)
						// fmt.Println(line)
						excelname[i] = line
						i++
						if err != nil {
							if err == io.EOF {
								fmt.Println("File read ok!")
								break
							} else {
								fmt.Println("Read file error!", err)
								return
							}
						}
					}

					fmt.Println(excelname)
					bool1 := true
					for i = 2; i < 7; i++ {
						bool2, _ := excel.PathExists("./" + excelname[i])
						if bool2 == false {
							fmt.Println(excelname[i])
							bool1 = false
							break
						}
					}

					if bool1 {
						go func() {
							excel.Main(excelname)
							outTE.SetText("运行完成")
							isSpecialMode.SetSatisfied(true)
						}()

						openFile, e := os.OpenFile("./name.txt", os.O_RDWR|os.O_CREATE|os.O_TRUNC, 777)
						if e != nil {
							fmt.Println(e)
						}
						openFile.WriteString(inTE.Text())
						openFile.Close()
						// inbutthon.SetChecked(false)
						isSpecialMode.SetSatisfied(false)

						outTE.SetText("等着")
					} else {
						// isSpecialMode.SetSatisfied(false)
						// inbutthon.SetChecked(false)
						fmt.Println(excelname)
						outTE.SetText("文件名不对")
					}

				},
			},
			// PushButton{
			// 	Text: "点这里1",
			// 	// AssignTo: &inbutthon,
			// 	Enabled: Bind("isSpecialMode"),
			// 	OnClicked: func() {

			// 	},
			// },
		},
	}.Run()
}
