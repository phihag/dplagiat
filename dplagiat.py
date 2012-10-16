#!/usr/bin/env python
# coding: utf-8

import hashlib
import shutil
import io
import os.path
import optparse
import xml.dom.minidom
import zipfile

from lxml import etree

_DOCX_NAMESPACES = {
	'dc': 'http://purl.org/dc/elements/1.1/',
	'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
	'dcterms': 'http://purl.org/dc/terms/',
	'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
	'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

def main():
	parser = optparse.OptionParser('%prog file [file..]')
	parser.add_option('-i', '--handle-images', action='store_true')
	parser.add_option('--extract-images', action='store_true')
	parser.add_option('--extract-dir', metavar='DIR', default='./imgs')
	opts,args = parser.parse_args()

	if not args:
		parser.error('Expected at least one file as argument')
	for fn in args:
		analyze(fn, opts)
		print('\n\n')

def _xpath_text(root, xpath):
	nodes = root.xpath(xpath, namespaces=_DOCX_NAMESPACES)
	assert len(nodes) <= 1
	if len(nodes) == 0:
		return None
	return ''.join(nodes[0].xpath('text()'))

def docx_properties(zf, opts):
	bad_fn = zf.testzip()
	if bad_fn:
		raise ValueError('Not a docx file; zip error in ' + bad_fn)

	files = set(zf.namelist())

	image_files = [fn for fn in files if fn.startswith('word/media/image')]
	if opts.handle_images:
		from PIL import Image
		from PIL.ExifTags import TAGS
		def _get_exif(img):
			bio = io.BytesIO(img)
			i = Image.open(bio)
			return {TAGS.get(tag, tag): value for tag,value in i._getexif().items()}

		images = {}
		for fn in image_files:
			img = zf.read(fn)
			csum = hashlib.sha256(img).hexdigest()
			if csum in images:
				images[csum]['filenames'].add(fn)
			else:
				images[csum] = {
					'filenames': set([fn]),
					'exif': _get_exif(img),
					'content': img,
				}
		if images:
			print(str(len(images)) + ' ' + ('Bild' if len(images) == 1 else 'Bilder') + ': ' + ', '.join(csum + ' (' + idata['exif']['Software'] + ')' for csum, idata in images.items()))
			if opts.extract_images:
				for csum,idata in images.items():
					if not os.path.exists(opts.extract_dir):
						os.makedirs(opts.extract_dir)
					imgfn = os.path.join(opts.extract_dir, csum + os.path.splitext(idata['filenames'][0])[1])
					with open(imgfn, 'wb') as imgf:
						imgf.write(idata['content'])
						print('Extracted ' + imgfn)
	files.difference_update(image_files)

	settings = etree.fromstring(zf.read('word/settings.xml'))
	res = {}
	res['revisions_index'] = len(settings.xpath('//w:rsid', namespaces=_DOCX_NAMESPACES))
	files.remove('word/settings.xml')

	props_app = etree.fromstring(zf.read('docProps/app.xml'))
	template = _xpath_text(props_app, '/ep:Properties/ep:Template')
	app = _xpath_text(props_app, '/ep:Properties/ep:Application')
	appVersion = _xpath_text(props_app, '/ep:Properties/ep:AppVersion')
	if appVersion is None:
		appId = '(Offiziell) ' + app + ' (wahrscheinlich LibreOffice/OpenOffice)'
	else:
		appVersion = {
			'14.0000': '2010',
		}.get(appVersion, appVersion)
		appId = app + ' ' + appVersion
	print(appId + ' (Template: ' + template + ')')
	files.remove('docProps/app.xml')

	props_core = etree.fromstring(zf.read('docProps/core.xml'))
	creator = _xpath_text(props_core, '//dc:creator')
	lastModifiedBy = _xpath_text(props_core, '//cp:lastModifiedBy')
	print('Autor: ' + creator + (' (Zuletzt geändert von ' + lastModifiedBy + ')' if creator != lastModifiedBy else ''))
	title = _xpath_text(props_core, '//dc:title')
	if title is not None:
		print('Titel: ' + (title if title else '[leer]'))
	res['created'] = _xpath_text(props_core, '//dcterms:created')
	res['modified'] = _xpath_text(props_core, '//dcterms:modified')
	lastPrinted = _xpath_text(props_core, '//cp:lastPrinted')
	if lastPrinted is not None:
		res['last_printed'] = lastPrinted
	res['revisions_metadata'] = int(_xpath_text(props_core, '//cp:revision'))
	files.remove('docProps/core.xml')

	print('Erstellt: ' + res['created'] + ', zuletzt bearbeitet: ' + res['modified'] + (', Zuletzt gedruckt: ' + res['last_printed'] if 'last_printed' in res else ''))
	print('Text-Revisionen: ' + str(res['revisions_index']) + ', Dokument-Revisionen: ' + str(res['revisions_metadata']-1))

	files.discard('[Content_Types].xml') # machine-generated and content-free
	files.discard('_rels/.rels') # always identical
	files.discard('word/_rels/footnotes.xml.rels') # Automatically generated based on text
	# print(files) # Remaining files are ignored for now

def analyze(fn, opts):
	print(os.path.basename(fn))
	with open(fn, 'rb') as f:
		print (hashlib.sha256(f.read()).hexdigest())
	with zipfile.ZipFile(fn) as zf:
		ps = docx_properties(zf, opts)

if __name__ == '__main__':
	main()

# TODO: fonts
