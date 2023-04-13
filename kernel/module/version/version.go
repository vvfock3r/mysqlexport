package version

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

const AppVersion = "v0.0.1"

// Version implement the Module interface
type Version struct {
	flag bool
}

func (v *Version) Register(cmd *cobra.Command) {
	// register flag -v / --version
	cmd.PersistentFlags().BoolVarP(&v.flag, "version", "v", false, "version message")
}

func (v *Version) MustCheck(*cobra.Command) {
	if v.flag {
		fmt.Println(AppVersion)
		os.Exit(0)
	}
}

func (v *Version) Initialize(*cobra.Command) error {
	return nil
}
