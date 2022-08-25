import * as vscode from 'vscode';

export function activate(context: vscode.ExtensionContext) {
	// set context keys
	vscode.commands.executeCommand('setContext', 'vba.fileExts', ['.docm', '.xlsm'])

	// register commands
	const commands = {
		'vba.extract': () => {
		},
		'vba.write': () => {
		},
	}
	Object.entries(commands).forEach(([name, callback]) => {
		vscode.commands.registerCommand(name, callback)
	})

	console.log('vscode-vba activated.')
}
