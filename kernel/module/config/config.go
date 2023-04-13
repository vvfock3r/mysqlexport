package config

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
	"github.com/spf13/viper"
)

// Config implement the Module interface
type Config struct {
	flag      string
	AddFlag   bool
	Name      string
	Exts      []string
	Path      []string
	MustExist bool
}

func (c *Config) Register(cmd *cobra.Command) {
	// register flag -c / --config
	if c.AddFlag {
		cmd.PersistentFlags().StringVarP(&c.flag, "config", "c", "", "config file")
	}

	// supported config extensions
	viper.SupportedExts = c.Exts

	// search config file
	viper.SetConfigName(c.Name)
	for _, path := range c.Path {
		viper.AddConfigPath(path)
	}
}

func (c *Config) MustCheck(*cobra.Command) {
	// if -c / --config is specified, the file must exist
	if c.flag != "" {
		_, err := os.Stat(c.flag)
		if err != nil && os.IsNotExist(err) {
			fmt.Printf("cannot find the file: %s\n", c.flag)
			os.Exit(1)
		}
	}
}

func (c *Config) Initialize(*cobra.Command) error {
	// read configuration
	viper.SetConfigFile(c.flag)
	err := viper.ReadInConfig()
	if err == nil {
		return nil
	}

	// if MustExist is set to false, ignore ConfigFileNotFoundError
	_, ok := err.(viper.ConfigFileNotFoundError)
	if ok && !c.MustExist {
		return nil
	}

	return err
}
