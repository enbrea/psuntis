# Copyright (c) STÜBER SYSTEMS GmbH. All rights reserved.
# Licensed under the MIT License.

BeforeAll {
	# Get the path of our module 
	$modulePath = $PSCommandPath.Replace('.tests.ps1','.psm1')
	# Import the module for testing
	Import-Module $modulePath -Force
}

Describe -name "Tests" {
	Context "Module interface" {
		It "Module should export 4 commands in alphabetical order." {
			$commands = Get-Command -Module PsUntis
			$commands.Count | Should -BeExactly 4
			$commands[0].Name | Should -Be "Get-UntisTerms"
			$commands[1].Name | Should -Be "Set-UntisTermAsActive"
			$commands[2].Name | Should -Be "Start-UntisBackup"
			$commands[3].Name | Should -Be "Start-UntisExport"
		}
	}
}