package main

import (
	"fmt"
	//"path/filepath"
	"strings"

	"github.com/Luxurioust/excelize"
	"os"
	"sync"
)

var (
	excelPath string
	luaPath string
	workResultLock sync.WaitGroup
)

func main() {
	//读取配置文件
	configExcel, configErr := excelize.OpenFile("config.xlsx")
	if configErr != nil {
		fmt.Println("read config.xlsx err==========")
		return
	}

	rows := configExcel.GetRows("sheet1")
	if len(rows) < 3 {
		fmt.Println("config.xlsx格式要求至少要三行数据，否则认为无效")
		return
	}

	excelPath = rows[0][1]
	luaPath = rows[1][1]
	fmt.Println("excelPath is ", excelPath, " luaPath is ", luaPath)

	fmt.Println("begin to work ............")
	for k, row := range rows {
		fmt.Println("loop =============1  ", k)
		if k < 2 {
			fmt.Println("loop =============2  ", k)
			continue
		}
		fmt.Println("loop =============3  ", k)
		if row[0] == "" || row[1] == "" {
			fmt.Println("config break at row ", k)
			break
		}
		fmt.Println("loop =============4  ", k, row[0], row[1])
		go writeLuaFile(row[0], row[1])

	}

	workResultLock.Wait()
	fmt.Println("work finished ............")


	// 枚举配置目录的所有文件
	//filepath.Walk("./", walkCallback)
}

/*func walkCallback(path string, f os.FileInfo, err error) error {

	go writeLuaFile(f, err)

	return err
}*/

func writeLuaFile(excelFileNnme string, luaFilename string) error{
	workResultLock.Add(1)
	if excelFileNnme == "" || luaFilename == "" {
		fmt.Println("excelFileNnme luaFilename =", excelFileNnme, luaFilename)
		return nil
	}

	fmt.Println("file name is " + excelFileNnme)
	var luaStr string = ""
	xlsx, openErr := excelize.OpenFile(excelPath + excelFileNnme)
	if openErr != nil {
		fmt.Println(openErr)
		return openErr
	}

	rows := xlsx.GetRows("sheet1")
	if len(rows) < 5 {
		fmt.Println("数据表格式要求至少要五行数据，否则认为是空表")
		return nil
	}

	// 写注释 标明每个字段的含义
	luaStr += "-- "
	for i := 0; i < len(rows[1]); i++ {
		if rows[1][i] == "" {
			break
		}
		luaStr += rows[1][i] + ":" + rows[0][i] + ", "
	}
	luaStr += "\n"

	//开始写lua配置表
	luaStr += "local " + luaFilename + " = {\n"
	for k, row := range rows {
		if k < 4 {
			continue
		}
		// 遇到第一个元素为空的行 直接跳出循环
		if row[0] == "" {
			break
		}

		luaStr += "\t{"

		// 可以考虑给string类型加上 ""
		for cellKey, colCell := range row {
			if rows[1][cellKey] == "" || colCell == "" {
				continue
			}

			// 判断是否是string类型
			cellStr := ""
			if (strings.Contains(rows[2][cellKey], "string")) {
				if strings.Contains(rows[2][cellKey], "[]") {
					strArray := strings.Split(colCell, ",")
					for _, tempStr := range strArray {
						cellStr += "'" + tempStr + "'" + ","
					}
				}else {
					cellStr = "'" + colCell + "'"
				}
			}else {
				cellStr = colCell
			}


			if strings.Contains(rows[2][cellKey], "[]") {
				luaStr += rows[1][cellKey] + " = {" + cellStr + "}, "
			}else {
				luaStr += rows[1][cellKey] + " = " + cellStr + ", "
			}
		}

		luaStr += "};\n"
	}
	luaStr += "}\n"
	luaStr += "return " + luaFilename

	// 写入到配置的目录中
	luaFile, fileErr := os.Create(luaPath + luaFilename + ".lua" )
	if fileErr != nil {
		fmt.Println(fileErr)
		return fileErr
	}
	defer luaFile.Close()
	_, writeErr := luaFile.WriteString(luaStr)
	if writeErr != nil {
		fmt.Println(writeErr)
		return writeErr
	}
	luaFile.Sync()
	workResultLock.Done()

	return  nil
}
