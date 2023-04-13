package watch

import (
	"fmt"
	"path/filepath"
	"strings"

	"github.com/fsnotify/fsnotify"
	"github.com/spf13/cobra"
	"github.com/spf13/viper"
	"go.uber.org/zap"

	"github.com/vvfock3r/mysqlexport/kernel/iface"
	"github.com/vvfock3r/mysqlexport/kernel/module/logger"
)

// Watch implement the Module interface
type Watch struct {
	List []iface.Module
}

func (w *Watch) Register(*cobra.Command) {}

func (w *Watch) MustCheck(*cobra.Command) {}

func (w *Watch) Initialize(cmd *cobra.Command) error {
	// skip if config file is not used
	if viper.ConfigFileUsed() == "" {
		return nil
	}

	// config watch
	viper.WatchConfig()
	viper.OnConfigChange(func(e fsnotify.Event) {
		// print log
		fileName := e.Name
		fileAbsName, err := filepath.Abs(fileName)
		if err == nil {
			fileName = fileAbsName
		}
		fileName = filepath.ToSlash(fileName)
		operation := strings.ToLower(e.Op.String())
		logger.Warn("config update trigger",
			zap.String("operation", operation),
			zap.String("filename", fileName))

		// initialize
		for _, m := range w.List {
			err = m.Initialize(cmd)
			if err != nil {
				logger.Warn("config reload ignored",
					zap.String("object", fmt.Sprintf("%T", m)),
					zap.String("detail", err.Error()))
			} else {
				logger.Warn("config reload success",
					zap.String("object", fmt.Sprintf("%T", m)),
					zap.String("detail", "success"))
			}
		}
	})
	return nil
}
