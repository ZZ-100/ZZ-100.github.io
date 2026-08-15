/**
 * 从 source/about/index.md 的 Contact 部分解析联系方式，
 * 供侧边栏 Links 图标使用（与 About 页 Contact 内容自动同步）。
 * 格式：- 名称: [显示文字](链接) 或 - 名称: https://... 或 - 名称: xxx@yyy.com
 */
var fs = require('fs')
var path = require('path')

var ICONS = {
    'email': 'fas fa-envelope',
    'mail': 'fas fa-envelope',
    'e-mail': 'fas fa-envelope',
    'github': 'fab fa-github',
    'google scholar': 'fas fa-graduation-cap',
    'scholar': 'fas fa-graduation-cap',
    'orcid': 'fab fa-orcid',
    'linkedin': 'fab fa-linkedin',
    'researchgate': 'fab fa-researchgate',
    'twitter': 'fab fa-twitter',
    'x': 'fab fa-x-twitter',
    'facebook': 'fab fa-facebook',
    'instagram': 'fab fa-instagram',
    'youtube': 'fab fa-youtube',
    'wechat': 'fab fa-weixin',
    'weixin': 'fab fa-weixin',
    'qq': 'fab fa-qq',
    'bilibili': 'fab fa-bilibili',
    'homepage': 'fas fa-home',
    'website': 'fas fa-globe',
    'blog': 'fas fa-blog',
    'phone': 'fas fa-phone',
    'tel': 'fas fa-phone',
    'address': 'fas fa-map-marker-alt',
    'university': 'fas fa-university',
    'school': 'fas fa-school'
}

function iconFor(name) {
    var key = name.toLowerCase().trim()
        .replace(/^personal\s+/, '')
        .replace(/\s*\(.*?\)\s*$/, '')
        .replace(/[-_ ]+\d+$/, '')
    if (ICONS[key]) return ICONS[key]
    if (/^mailto:/i.test(name)) return 'fas fa-envelope'
    return 'fas fa-link'
}

hexo.extend.helper.register('contact_items', function () {
    var aboutFile = path.join(hexo.source_dir, 'about', 'index.md')
    var content
    try {
        content = fs.readFileSync(aboutFile, 'utf8')
    } catch (e) {
        return []
    }
    var m = content.match(/##\s*Contact([\s\S]*?)(?=\n##\s|\s*$)/)
    if (!m) return []
    var section = m[1]
    var items = []
    section.split(/\r?\n/).forEach(function (line) {
        line = line.trim()
        if (!line) return
        var lm = line.match(/^[-*]\s+(.+?)\s*:\s*(.+)$/)
        if (!lm) return
        var name = lm[1].trim()
        var rest = lm[2].trim()
        var url = '', label = ''
        var link = rest.match(/^\[([^\]]*)\]\(([^)]+)\)$/)
        if (link) {
            label = link[1]
            url = link[2]
        } else if (/^https?:\/\/\S+$/.test(rest)) {
            url = rest
            label = rest.replace(/^https?:\/\//, '').replace(/\/$/, '')
        } else if (/^\S+@\S+\.\S+$/.test(rest)) {
            url = 'mailto:' + rest
            label = rest
        } else {
            return
        }
        items.push({ name: name, url: url, label: label || name, icon: iconFor(name) })
    })
    return items
})
