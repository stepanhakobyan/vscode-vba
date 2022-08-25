import * as fs from 'fs'
import * as vscode from 'vscode'
import * as yauzl from 'yauzl'

export function activate(context: vscode.ExtensionContext) {
	// set context keys
	vscode.commands.executeCommand('setContext', 'vba.fileExts', ['.docm', '.xlsm'])

	// register commands
	const commands = {
		'vba.extract': (uri: vscode.Uri) => {
			yauzl.open(uri.path, { lazyEntries: true }, (err, zipfile) => {
				if (err) {
					vscode.window.showErrorMessage(err.message)
					throw err
				}
				zipfile.readEntry()
				zipfile.on('entry', (entry: yauzl.Entry) => {
					if (entry.fileName == 'xl/vbaProject.bin') {
						zipfile.openReadStream(entry, (err, stream) => {
							if (err) {
								vscode.window.showErrorMessage(err.message)
								throw err
							}
							// TODO read vbaProject.bin
							const chunks: Buffer[] = []
							stream.on('data', (chunk) => {
								chunks.push(chunk)
							});
							stream.once('end', () => {
								// TODO read and convert ole to text
								const bufstr = Buffer.concat(chunks).toString()
								// TODO split text per module and output to files
							})
						})
						zipfile.close()
					} else {
						zipfile.readEntry()
					}
				})
			})
		},
		'vba.write': (uri: vscode.Uri) => {
			// TODO
		},
	}
	Object.entries(commands).forEach(([name, callback]) => {
		vscode.commands.registerCommand(name, callback)
	})

	console.log('vscode-vba activated.')
}
