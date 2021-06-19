export default function initSharedStrings() {
	const sharedStrings = []
	const sharedStringsIndex = {}

	return {
		getSharedStringsXml() {
			return generateXml(sharedStrings)
		},

		getSharedString(string) {
      let id = sharedStringsIndex[string]
      if (id === undefined) {
				id = String(sharedStrings.length)
				sharedStringsIndex[string] = id
				sharedStrings.push(string)
      }
      return id
		}
	}
}

function generateXml(sharedStrings) {
	let xml = '<?xml version="1.0"?>'
	xml += '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
	for (const string of sharedStrings) {
		xml += `<si><t>${string}</t></si>`
	}
	xml += '</sst>'
	return xml
}