bill-windows:
	 GOOS=windows GOARCH=amd64 go build  -o build/billcheck.exe github.com/jinlingan/billcheck/cmd/billcheck
run-bill:
	go run ./cmd/billcheck