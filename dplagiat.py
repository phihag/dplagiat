#!/usr/bin/env python
# coding: utf-8

from __future__ import division
from __future__ import print_function

import collections
import colorsys
import hashlib
import shutil
import io
import itertools
import os.path
import operator
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
	'wd': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
	'wd2010': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
}

def main():
	parser = optparse.OptionParser('%prog file [file..]')
	parser.add_option('-i', '--handle-images', action='store_true')
	parser.add_option('--extract-images', action='store_true')
	parser.add_option('-r', '--write-revision-html', action='store_true')
	parser.add_option('--extract-dir', metavar='DIR', default='./extracted')
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

def docx_docRevisions(doc):
	res = []

	revision = None
	namespace = '{' + _DOCX_NAMESPACES['w'] + '}'
	DEFAULT_ATTRIB = namespace + 'rsidRDefault'
	REV_ATTRIB = namespace + 'rsidR'
	def _visit_node(node, revision):
		if DEFAULT_ATTRIB in node.attrib:
			revision = node.attrib[DEFAULT_ATTRIB]
		if REV_ATTRIB in node.attrib:
			revision = node.attrib[REV_ATTRIB]

		if node.tag in [
					namespace + 'instrText', # Automatic reference
					'{' + _DOCX_NAMESPACES['wd'] + '}posOffset', # Drawing metadata
					'{' + _DOCX_NAMESPACES['wd2010'] + '}pctHeight', # Drawing metadata
					'{' + _DOCX_NAMESPACES['wd2010'] + '}pctWidth', # Drawing metadata
		]:
			return
		if node.text:
			res.append((node.text, revision))
		for child in node:
			_visit_node(child, revision)
		if node.tag == namespace + 'p':
			res.append(('\n', revision))
		if node.tail:
			res.append((node.tail, revision))

	for rootNode in doc.xpath('//w:body/*', namespaces=_DOCX_NAMESPACES):
		_visit_node(rootNode, None)
	return res

def _colors(alpha=0.8):
	# Algorithm from http://ridiculousfish.com/blog/posts/colors.html
	for idx in itertools.count():
		numbits = 32
		b = sum(1<<(numbits-1-i) for i in range(numbits) if idx>>i&1) / 2**32
		hue = (b + .6) % 1
		r,g,b = colorsys.hls_to_rgb(hue, .6, 1)
		yield 'rgba(%s, %s, %s, %s)' % (int(r*255),int(g*255),int(b*255),alpha)

def _revisionHTML(revData, docData):
	out = etree.fromstring(
"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<title></title>
<script type="text/javascript">
"use strict";

function highlightRev(rev) {
	var textContainer = document.getElementById("text");
	for (var i = 0; i != textContainer.childNodes.length;i++) {
		var textNode = textContainer.childNodes[i];
		var tnRev = textNode.getAttribute("data-rev");
		var cn = "rev-" + tnRev;
		if (typeof rev !== 'undefined') { // No AND because of required ampersand
			if (tnRev != rev) {
				cn += " text-shaded";
			}
		}
		if (textNode.className != cn) {
			textNode.className = cn;
		}
	}
}

function clickRev(ev) {
	highlightRev(ev.target.getAttribute('data-rev'));
}

function onLoad() {
	var revsNode = document.getElementById("revs");
	for (var i = 0;i != revsNode.childNodes.length;i++) {
		var revNode = revsNode.childNodes[i];
		revNode.addEventListener('click', clickRev, false);
	}

	var body = document.getElementsByTagName("body")[0];
	body.addEventListener('keydown', function(ev) {
		if (ev.keyCode == 27) {
			highlightRev(undefined);
		}
	}, false);
}


document.addEventListener('DOMContentLoaded', onLoad, false);
</script>
<style type="text/css">
#checksum,#template {font-family: monospace;}
#revs {width: 15em; position: absolute; position: fixed; right: 0; top: 0; height: 100%; overflow-y: auto;}
#revs a {display: block; color: black !important; text-decoration: none !important; font-size: 110%; padding: 0.5em 0em 0.5em 0.65em; font-family: sans-serif; }
#text {white-space: pre-wrap; margin-right: 17em;}
.text-shaded {opacity: 0.9;}
</style>
<style type="text/css" id="revStyle">

</style>
</head>
<body>
<h1></h1>
<p>
<div>SHA256: <span id="checksum"></span></div>
<div>Anzahl Speichervorgänge: <strong id="saveCount"></strong></div>
<div>Autor: <span id="creator"></span> (Zuletzt geändert von <span id="lastModifiedBy"></span>)</div>
<div id="titleContainer"></div>
<div>Erstellt: <span id="created"/>, zuletzt bearbeitet: <span id="modified"/>, zuletzt gedruckt: <span id="last_printed"/></div>
<div><span id="appId"></span> (<span id="template"></span>)</div>
</p>

<div id="revs"/>
<div id="text"></div>
</body>
</html>
""")

	out.xpath('//h1')[0].text = docData['filename']
	out.xpath('//title')[0].text = docData['filename']
	out.xpath('//*[@id="checksum"]')[0].text = docData['sha256']
	out.xpath('//*[@id="saveCount"]')[0].text = str(docData['revisions_metadata'])
	out.xpath('//*[@id="appId"]')[0].text = docData['appId']
	out.xpath('//*[@id="template"]')[0].text = docData['template']
	out.xpath('//*[@id="creator"]')[0].text = docData['creator']
	out.xpath('//*[@id="lastModifiedBy"]')[0].text = docData['lastModifiedBy']
	if 'title' in docData:
		titleContainer = out.xpath('//*[@id="titleContainer"]')[0]
		titleContainer.text = 'Titel: '
		titleNode = etree.Element('span')
		titleNode.text = docData['title']
		titleContainer.append(titleNode)
	out.xpath('//*[@id="created"]')[0].text = docData['created']
	out.xpath('//*[@id="modified"]')[0].text = docData['modified']
	if 'last_printed' in docData:
		out.xpath('//*[@id="last_printed"]')[0].text = docData['last_printed']

	revOrder = []
	seenRev = set()
	textNode = out.xpath('//*[@id="text"]')[0]
	for text,rev in revData:
		span = etree.Element('span')
		span.text = text
		span.attrib['class'] = 'rev-' + rev
		span.attrib['data-rev'] = rev
		if rev not in seenRev:
			revOrder.append(rev)
			span.attrib['id'] = 'first-' + rev
			seenRev.add(rev)
		textNode.append(span)

	revBytes = collections.defaultdict(int)
	for text,rev in revData:
		revBytes[rev] += len(text)
	byteCount = sum(revBytes.values())
	revsNode = out.xpath('//*[@id="revs"]')[0]
	for rev in revOrder:
		revNode = etree.Element('a')
		revNode.text = rev + ' (' + str(revBytes[rev]) + ' Zeichen)'
		revNode.attrib['title'] = str(int(100 * revBytes[rev] / byteCount)) + '%'
		revNode.attrib['href'] = '#first-' + rev
		revNode.attrib['class'] = 'rev-' + rev
		revNode.attrib['data-rev'] = rev
		revsNode.append(revNode)

	revStyleNode = out.xpath('//style[@id="revStyle"]')[0]
	for rev,color,shadedColor in zip(revOrder, _colors(0.7), _colors(0.2)):
		revStyleNode.text += '.rev-%s {background-color: %s;}\n' % (rev,color)
		revStyleNode.text += '.rev-%s.text-shaded {background-color: %s;}\n' % (rev,shadedColor)

	return etree.tostring(out)

def docx_properties(zf, filename, opts):
	res = {}
	res['filename'] = os.path.basename(filename)
	with open(filename, 'rb') as f:
		res['sha256'] = hashlib.sha256(f.read()).hexdigest()

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
	res['revisions_index'] = len(settings.xpath('//w:rsid', namespaces=_DOCX_NAMESPACES))
	files.remove('word/settings.xml')

	props_app = etree.fromstring(zf.read('docProps/app.xml'))
	res['template'] = _xpath_text(props_app, '/ep:Properties/ep:Template')
	app = _xpath_text(props_app, '/ep:Properties/ep:Application')
	appVersion = _xpath_text(props_app, '/ep:Properties/ep:AppVersion')
	if appVersion is None:
		res['appId'] = '(Offiziell) ' + app + ' (wahrscheinlich LibreOffice/OpenOffice)'
	else:
		appVersion = {
			'14.0000': '2010',
		}.get(appVersion, appVersion)
		res['appId'] = app + ' ' + appVersion
	files.remove('docProps/app.xml')

	props_core = etree.fromstring(zf.read('docProps/core.xml'))
	res['creator'] = _xpath_text(props_core, '//dc:creator')
	res['lastModifiedBy'] = _xpath_text(props_core, '//cp:lastModifiedBy')
	title = _xpath_text(props_core, '//dc:title')
	if title is not None:
		res['title'] = title if title else '[leer]'

	res['created'] = _xpath_text(props_core, '//dcterms:created')
	res['modified'] = _xpath_text(props_core, '//dcterms:modified')
	lastPrinted = _xpath_text(props_core, '//cp:lastPrinted')
	if lastPrinted is not None:
		res['last_printed'] = lastPrinted
	res['revisions_metadata'] = int(_xpath_text(props_core, '//cp:revision'))
	files.remove('docProps/core.xml')

	doc = etree.fromstring(zf.read('word/document.xml'))
	revData = docx_docRevisions(doc)
	if opts.write_revision_html:
		html = _revisionHTML(revData, res)
		if not os.path.exists(opts.extract_dir):
			os.makedirs(opts.extract_dir)
		with open(os.path.join(opts.extract_dir, res['filename'] + '.html'), 'wb') as outf:
			outf.write(html.encode('utf-8'))

	print(res['filename'])
	print(res['sha256'])
	print(res['appId'] + ' (Template: ' + res['template'] + ')')
	print('Autor: ' + res['creator'] + (' (Zuletzt geändert von ' + res['lastModifiedBy'] + ')' if res['creator'] != res['lastModifiedBy'] else ''))
	if 'title' in res:
		print('Titel: ' + res['title'])
	print('Erstellt: ' + res['created'] + ', zuletzt bearbeitet: ' + res['modified'] + (', Zuletzt gedruckt: ' + res['last_printed'] if 'last_printed' in res else ''))
	print('Text-Revisionen: ' + str(res['revisions_index']) + ', Dokument-Revisionen: ' + str(res['revisions_metadata']))

	files.discard('[Content_Types].xml') # machine-generated and content-free
	files.discard('_rels/.rels') # always identical
	files.discard('word/_rels/footnotes.xml.rels') # Automatically generated based on text
	# print(files) # Remaining files are ignored for now

def analyze(fn, opts):
	with zipfile.ZipFile(fn) as zf:
		ps = docx_properties(zf, fn, opts)

if __name__ == '__main__':
	main()

# TODO: fonts
