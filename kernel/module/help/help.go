package help

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

// Help implement the Module interface
type Help struct {
	HiddenShortFlag   bool
	HiddenHelpCommand bool
}

func (h *Help) Register(cmd *cobra.Command) {
	cmd.SetHelpFunc(func(command *cobra.Command, strings []string) {
		fmt.Println(`
Export mysql to excel                                               
For details, please refer to https://github.com/vvfock3r/mysqlexport

Usage:                                                                            
  mysqlexport [flags]                                                             

General Flags:
  -v, --version                     version message
      --help                        help message
	  
Log Flags:      
      --log-format string           log format (default "console")                
      --log-level string            log level (default "info")                    
      --log-output string           log output (default "stdout") 

MySQL Flags:
  -h, --host string                 mysql host (default "127.0.0.1")
  -P, --port string                 mysql port (default "3306")  
  -u, --user string                 mysql user (default "root")  
  -p, --password string             mysql password
  -d, --database string             mysql database
  -e, --execute string              execute sql command                             
      --charset string              mysql charset (default "utf8mb4")    
      --collation string            mysql collation (default "utf8mb4_general_ci")	  
      --connect-timeout string      mysql connect timeout (default "5s")       
      --read-timeout string         mysql read timeout (default "30s")	  
      --write-timeout string        mysql write timeout (default "30s")	  
      --max-allowed-packet string   mysql max allowed packet (default "16MB")   	  
      --batch-size int              batch size (default 10000)                      
      --sleep-time string           sleep time (default "1s")

Excel Flags:
  -o, --output string               output xlsx file
      --excel-password string       excel-password                                
      --col-align string            col align (default "left")
      --col-width string            col-width
      --row-height string           row height
      --sheet-line int              max line per sheet (default 1000000)
      --sheet-name string           sheet name (default "Sheet")`)
		os.Exit(0)
	})

	if h.HiddenShortFlag {
		cmd.PersistentFlags().Bool("help", false, "help message")
	} else {
		cmd.PersistentFlags().BoolP("help", "h", false, "help message")
	}

	if h.HiddenHelpCommand {
		cmd.SetHelpCommand(&cobra.Command{Use: "no-help", Hidden: true})
	}
}

func (h *Help) MustCheck(*cobra.Command) {}

func (h *Help) Initialize(*cobra.Command) error {
	return nil
}
