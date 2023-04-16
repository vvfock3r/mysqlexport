package version

import (
	"fmt"
	"os"
	"runtime"

	"github.com/spf13/cobra"
)

const AppVersion = "v1.1.0"

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
		fmt.Printf("mysqlexport version %s %s/%s\n", AppVersion, runtime.GOOS, runtime.GOARCH)
		os.Exit(0)
	}
}

func (v *Version) Initialize(*cobra.Command) error {
	return nil
}
