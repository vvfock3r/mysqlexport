package help

import (
	"fmt"
	"github.com/spf13/cobra"
	"os"
)

// Help implement the Module interface
type Help struct {
	HiddenShortFlag   bool
	HiddenHelpCommand bool
}

func (h *Help) Register(cmd *cobra.Command) {
	cmd.SetHelpFunc(func(command *cobra.Command, strings []string) {
		var msg = `
Export mysql to excel                                               
For details, please refer to https://github.com/vvfock3r/mysqlexport

Usage:                                                                                                         
  mysqlexport [flags]                                                                                          
                                                                                                               
General Flags:
  -v, --version                     version message
      --help                        displays the help message for the program
	  
Log Flags:
      --log-level string            specifies the level of logging (default "info")
      --log-format string           specifies the log format (default "console")
      --log-output string           specifies the log output destination (default "stdout")

MySQL Flags:
  -h, --host string                 specifies the MySQL host (default "127.0.0.1")
  -P, --port string                 specifies the MySQL port (default "3306")
  -u, --user string                 specifies the MySQL user (default "root")
  -p, --password string             specifies the MySQL password
  -d, --database string             specifies the MySQL database
  -e, --execute string              specifies the SQL command to be executed
      --charset string              specifies the MySQL charset (default "utf8mb4")
      --collation string            specifies the MySQL collation (default "utf8mb4_general_ci")
      --connect-timeout string      specifies the MySQL connection timeout (default "5s")
      --read-timeout string         specifies the MySQL read timeout (default "30s")
      --write-timeout string        specifies the MySQL write timeout (default "30s")
      --max-allowed-packet string   specifies the MySQL maximum allowed packet (default "16MB")
      --batch-size int              specifies the batch size to use when executing SQL commands (default 10000)
      --delay-time string           specifies the time to delay between batches when executing SQL (default "1s")
	  
Excel Flags:
  -o, --output string               specifies the name of the output Excel file
      --setup-password string       specifies the password for the Excel file
      --sheet-name string           specifies the name of the sheet in the Excel file
      --sheet-line int              specifies the maximum number of lines per sheet in the Excel file (default 1000000)	  
      --row-height string           specifies the row height in the Excel file
      --col-width string            specifies the column width in the Excel file
      --col-align string            specifies the column alignment in the Excel file`
		fmt.Println(msg)
		os.Exit(0)
	})

	if h.HiddenShortFlag {
		cmd.PersistentFlags().Bool("help", false, "displays the help message for the program")
	} else {
		cmd.PersistentFlags().BoolP("help", "h", false, "displays the help message for the program")
	}

	if h.HiddenHelpCommand {
		cmd.SetHelpCommand(&cobra.Command{Use: "no-help", Hidden: true})
	}
}

func (h *Help) MustCheck(*cobra.Command) {}

func (h *Help) Initialize(*cobra.Command) error {
	return nil
}
