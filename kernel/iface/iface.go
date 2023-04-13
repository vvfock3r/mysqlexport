package iface

import "github.com/spf13/cobra"

type Module interface {
	Register(*cobra.Command)
	MustCheck(*cobra.Command)
	Initialize(*cobra.Command) error
}
