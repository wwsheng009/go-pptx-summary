package main

import (
	"archive/zip"
	"fmt"
	"io/ioutil"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

func main() {

	logfile, err := os.OpenFile("_summary_detail.txt", os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)
	if err != nil {
		fmt.Println(err)
		return
	}
	logfile2, err := os.OpenFile("_summary_title.txt", os.O_WRONLY|os.O_CREATE|os.O_TRUNC, 0644)

	if err != nil {
		fmt.Println(err)
		return
	}
	defer logfile.Close()
	defer logfile2.Close()

	re := regexp.MustCompile(`<a:t>(.*?)</a:t>`)

	dir, err := os.Getwd()
	if err != nil {
		fmt.Println(err)
	}
	fmt.Println("Current directory:", dir)

	// searchDir := "J:\\PPT分类\\PPT\\70套计划书模板\\模板"
	searchDir := dir
	err = filepath.Walk(searchDir, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			logfile.WriteString(err.Error() + "\n\n")
			logfile2.WriteString(err.Error() + "\n\n")
			return nil
		}
		if !info.IsDir() && strings.Contains(filepath.Ext(path), "pptx") {
			if strings.Contains(path, "~$") {
				return nil
			}

			logfile.WriteString(path + "\n\n")
			logfile2.WriteString(path + "\n\n")

			r, err := zip.OpenReader(path)
			if err != nil {
				logfile.WriteString(err.Error() + "\n\n")
				logfile2.WriteString(err.Error() + "\n\n")
				return nil
			}
			defer r.Close()

			filePath := "ppt/slides/slide"
			var file *zip.File

			var contentList [][2]string

			for _, f := range r.File {
				if strings.Contains(f.Name, filePath) {
					filename := filepath.Base(f.Name)
					fileExt := filepath.Ext(f.Name)
					slidename := filename[0 : len(filename)-len(fileExt)]

					file = f
					rc, err := file.Open()
					if err != nil {
						logfile.WriteString(err.Error() + "\n\n")
						logfile2.WriteString(err.Error() + "\n\n")
						return nil
					}
					content, err := ioutil.ReadAll(rc)
					if err != nil {
						logfile.WriteString(err.Error() + "\n\n")
						logfile2.WriteString(err.Error() + "\n\n")
						rc.Close()
						return nil
					}
					matches := re.FindAllStringSubmatch(string(content), -1)
					content1 := ""
					for _, v := range matches {
						xl := fixTitle(v[1])
						if xl != "" {
							content1 += xl + "-"
						}
					}
					if content1 != "" {
						content1 = strings.Trim(content1, "-")
						contentList = append(contentList, [2]string{slidename, content1})
					}
					rc.Close()
				}
			}
			// sort the array by the first element
			sort.Slice(contentList, func(i, j int) bool {
				title1 := strings.ReplaceAll(contentList[i][0], "slide", "")
				title2 := strings.ReplaceAll(contentList[j][0], "slide", "")

				t1, err := strconv.Atoi(title1)
				if err != nil {
					fmt.Println("Error:", err)
					return false
				}
				t2, err := strconv.Atoi(title2)
				if err != nil {
					fmt.Println("Error:", err)
					return false
				}
				return t1 < t2
			})
			write_title := false

			for _, v := range contentList {
				content := v[1]
				logfile.WriteString(v[0] + "\n\n")
				logfile.WriteString(content + "\n\n")
				if !write_title && content != "" {
					logfile2.WriteString(v[1] + "\n")
					write_title = true
				}
			}

			logfile.WriteString("\n")
			logfile2.WriteString("\n")
		}
		return nil
	})
	if err != nil {
		fmt.Println(err)
	}

	fmt.Println("Done")

}

func fixTitle(title string) string {
	lower := strings.ToLower(title)
	if strings.Contains(title, "标题") {
		return ""
	}
	if strings.Contains(lower, "title") {
		return ""
	}

	if strings.Contains(title, "描述") {
		return ""
	}
	if strings.Contains(lower, "description") {
		return ""
	}
	if title == "模板" {
		return ""
	}
	if lower == "template" {
		return ""
	}
	if strings.Contains(lower, "xx") {
		return ""
	}
	if strings.Contains(lower, "ppt") {
		return ""
	}
	if strings.Contains(title, "点击") {
		return ""
	}
	if strings.Contains(title, "输入") {
		return ""
	}
	if strings.Contains(title, "关键词") {
		return ""
	}
	if strings.Contains(title, "添加") {
		return ""
	}
	if strings.Contains(title, "目录") {
		return ""
	}
	if strings.Contains(lower, "click") {
		return ""
	}
	title = strings.ReplaceAll(title, "-", "-")
	title = strings.ReplaceAll(title, "·", "-")
	title = strings.ReplaceAll(title, "：", "-")
	title = strings.ReplaceAll(title, "。", "-")
	title = strings.ReplaceAll(title, "。", "-")
	title = strings.ReplaceAll(title, "/", "-")
	title = strings.ReplaceAll(title, " ", "")
	title = strings.ReplaceAll(title, "/", "-")
	title = strings.ReplaceAll(title, "--", "-")
	for strings.Contains(title, "--") {
		title = strings.ReplaceAll(title, "--", "-")
	}
	title = strings.Trim(title, "-")
	return title
}
