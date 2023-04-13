package load

import (
	"github.com/vvfock3r/mysqlexport/kernel/iface"
	"github.com/vvfock3r/mysqlexport/kernel/module/help"
	"github.com/vvfock3r/mysqlexport/kernel/module/logger"
	"github.com/vvfock3r/mysqlexport/kernel/module/mysql"
	"github.com/vvfock3r/mysqlexport/kernel/module/version"
)

// ModuleList 包含所有内置模块的列表
var ModuleList = []iface.Module{
	// 独立的模块放
	&version.Version{},
	&help.Help{
		HiddenShortFlag:   true,
		HiddenHelpCommand: true,
	},

	// 具有依赖关系的模块,详情可以查看模块的import部分
	&logger.Logger{
		AddFlag:   true,
		AddCaller: false,
	},
	&mysql.MySQL{
		AddFlag:         true,
		AllowedCommands: []string{"mysqlexport"},
	},
}
