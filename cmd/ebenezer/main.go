package main

import (
	"context"
	"log"
	"os"

	"ebenezer/internal/cmd"
)

func main() {
	ctx := context.Background()
	args := os.Args[1:]
	if err := cmd.Execute(ctx, args); err != nil {
		log.Fatal(err)
	}
}
