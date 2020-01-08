package utils

import (
	"bufio"
	"fmt"
	"os"
	"strings"

	log "github.com/sirupsen/logrus"
)

// 获取用户输入
func GetInput(out string) string {
	reader := bufio.NewReader(os.Stdin)
	fmt.Print(out + "\n")
	text, err := reader.ReadString('\n')
	text = strings.TrimSpace(text)
	if err != nil {
		log.Warnf("读取输入异常 %s", err)
	}
	return text
}
