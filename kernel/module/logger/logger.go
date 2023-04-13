package logger

import (
	"fmt"
	"os"
	"strings"
	"time"

	"github.com/spf13/cobra"
	"github.com/spf13/viper"
	"go.uber.org/zap"
	"go.uber.org/zap/zapcore"
)

// default logger
var DefaultLogger, _ = defaultLogConfig.build()

var defaultLogConfig = &LogConfig{
	Level:     "info",
	Format:    "console",
	Output:    "stdout",
	addCaller: true,
}

// Logger implement the Module interface
type Logger struct {
	AddFlag    bool
	AddCaller  bool
	Stacktrace zapcore.LevelEnabler
}

func (l *Logger) Register(cmd *cobra.Command) {
	var (
		defaultLogLevelKey  = "settings.log.level"
		defaultLogFormatKey = "settings.log.format"
		defaultLogOutputKey = "settings.log.output"
	)
	if !l.AddFlag {
		// default
		viper.SetDefault(defaultLogLevelKey, defaultLogConfig.Level)
		viper.SetDefault(defaultLogFormatKey, defaultLogConfig.Format)
		viper.SetDefault(defaultLogOutputKey, defaultLogConfig.Output)
		return
	}

	// flags
	cmd.PersistentFlags().String("log-level", defaultLogConfig.Level, "log level")
	cmd.PersistentFlags().String("log-format", defaultLogConfig.Format, "log format")
	cmd.PersistentFlags().String("log-output", defaultLogConfig.Output, "log output")

	// bind
	err := viper.BindPFlag(defaultLogLevelKey, cmd.PersistentFlags().Lookup("log-level"))
	if err != nil {
		panic(err)
	}
	err = viper.BindPFlag(defaultLogFormatKey, cmd.PersistentFlags().Lookup("log-format"))
	if err != nil {
		panic(err)
	}
	err = viper.BindPFlag(defaultLogOutputKey, cmd.PersistentFlags().Lookup("log-output"))
	if err != nil {
		panic(err)
	}
}

func (l *Logger) MustCheck(*cobra.Command) {}

func (l *Logger) Initialize(cmd *cobra.Command) error {
	logConfig := LogConfig{
		Level:      viper.GetString("settings.log.level"),
		Format:     viper.GetString("settings.log.format"),
		Output:     viper.GetString("settings.log.output"),
		addCaller:  l.AddCaller,
		stacktrace: l.Stacktrace,
	}

	newLogger, err := logConfig.build()
	if err != nil {
		return err
	}

	if DefaultLogger != nil {
		_ = DefaultLogger.Sync()
		for _, file := range defaultLogConfig.preOutputFiles {
			_ = file.Close()
		}
	}

	DefaultLogger = newLogger

	return nil
}

// LogConfig zap log config
type LogConfig struct {
	Level  string
	Format string
	Output string

	addCaller  bool
	stacktrace zapcore.LevelEnabler

	preOutputFiles []*os.File
	curOutputFiles []*os.File
}

func (l *LogConfig) build() (*zap.Logger, error) {
	level, err := l.level()
	if err != nil {
		return nil, err
	}

	encoder, err := l.encoder()
	if err != nil {
		return nil, err
	}

	syncers, err := l.output()
	if err != nil {
		return nil, err
	}

	logger := zap.New(zapcore.NewCore(encoder, syncers, level))

	if l.addCaller {
		logger = logger.WithOptions(zap.AddCaller(), zap.AddCallerSkip(1))
	}

	if l.stacktrace != nil {
		logger = logger.WithOptions(zap.AddStacktrace(l.stacktrace))
	}

	return logger, nil
}

func (l *LogConfig) level() (zap.AtomicLevel, error) {
	level, err := zapcore.ParseLevel(l.Level)
	if err != nil {
		unrecognized := "unrecognized log level: " + l.Level
		supported := "supported values: debug,info,warn,error,dpanic,panic,fatal"
		return zap.NewAtomicLevelAt(zapcore.InvalidLevel), fmt.Errorf(unrecognized + ", " + supported)
	}
	return zap.NewAtomicLevelAt(level), nil
}

func (l *LogConfig) encoder() (zapcore.Encoder, error) {
	// encoderConfig
	encoderConfig := zapcore.EncoderConfig{
		TimeKey:          "time",
		LevelKey:         "level",
		NameKey:          "logger",
		CallerKey:        "caller",
		FunctionKey:      zapcore.OmitKey,
		MessageKey:       "message",
		StacktraceKey:    "stacktrace",
		LineEnding:       zapcore.DefaultLineEnding,
		EncodeLevel:      zapcore.LowercaseLevelEncoder,
		EncodeTime:       zapcore.TimeEncoderOfLayout(time.DateTime),
		EncodeDuration:   zapcore.SecondsDurationEncoder,
		EncodeCaller:     zapcore.ShortCallerEncoder,
		EncodeName:       func(s string, encoder zapcore.PrimitiveArrayEncoder) { encoder.AppendString(s) },
		ConsoleSeparator: "",
	}

	// encoder
	switch l.Format {
	case "json":
		return zapcore.NewJSONEncoder(encoderConfig), nil
	case "console":
		encoderConfig.EncodeLevel = zapcore.CapitalLevelEncoder
		return zapcore.NewConsoleEncoder(encoderConfig), nil
	default:
		unrecognized := "unrecognized log format: " + l.Format
		supported := "supported values: json,console"
		return zapcore.NewConsoleEncoder(encoderConfig), fmt.Errorf(unrecognized + ", " + supported)
	}
}

func (l *LogConfig) output() (zapcore.WriteSyncer, error) {
	var (
		writeSyncers   []zapcore.WriteSyncer
		newOutputFiles []*os.File
	)

	for _, out := range strings.Split(l.Output, ",") {
		switch out {
		case "stdout":
			writeSyncers = append(writeSyncers, zapcore.AddSync(os.Stdout))
		case "stderr":
			writeSyncers = append(writeSyncers, zapcore.AddSync(os.Stderr))
		default:
			file, err := os.OpenFile(out, os.O_CREATE|os.O_APPEND|os.O_WRONLY, 0644)
			if err != nil {
				return nil, err
			}
			writeSyncers = append(writeSyncers, zapcore.Lock(zapcore.AddSync(file)))
			newOutputFiles = append(newOutputFiles, file)
		}
	}

	l.preOutputFiles = l.curOutputFiles
	l.curOutputFiles = newOutputFiles

	return zapcore.NewMultiWriteSyncer(writeSyncers...), nil
}

func Debug(msg string, fields ...zap.Field) {
	DefaultLogger.Debug(msg, fields...)
}

func Info(msg string, fields ...zap.Field) {
	DefaultLogger.Info(msg, fields...)
}

func Warn(msg string, fields ...zap.Field) {
	DefaultLogger.Warn(msg, fields...)
}

func Error(msg string, fields ...zap.Field) {
	DefaultLogger.Error(msg, fields...)
}

func DPanic(msg string, fields ...zap.Field) {
	DefaultLogger.DPanic(msg, fields...)
}

func Panic(msg string, fields ...zap.Field) {
	DefaultLogger.Panic(msg, fields...)
}

func Fatal(msg string, fields ...zap.Field) {
	DefaultLogger.Fatal(msg, fields...)
}
