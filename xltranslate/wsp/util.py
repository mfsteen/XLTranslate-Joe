#

def sanitise_string(text):
    text = u'%s' % (text, )
    text = text.strip()
    text = text.replace('\n', ' ')
    text = text.encode('ascii', 'ignore')
    return text
