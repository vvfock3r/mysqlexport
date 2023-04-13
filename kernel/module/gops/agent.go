package gops

import (
	"github.com/google/gops/agent"
	"github.com/spf13/cobra"
)

// Agent implement the Module interface
type Agent struct{}

func (a *Agent) Register(*cobra.Command) {}

func (a *Agent) MustCheck(*cobra.Command) {}

func (a *Agent) Initialize(*cobra.Command) error {
	return agent.Listen(agent.Options{})
}
