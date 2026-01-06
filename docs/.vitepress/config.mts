import { defineConfig } from 'vitepress'

// https://vitepress.dev/reference/site-config
export default defineConfig({
  title: "excelifyjs",
  description: "High‑performance Rust‑powered library for streaming, creating, and converting Excel .xlsx files in Node.js.",
  themeConfig: {
    // https://vitepress.dev/reference/default-theme-config
    nav: [
      { text: 'Home', link: '/' },
      { text: 'Guide', link: '/getting-started' },
      { text: 'Reference', link: '/workbook' }
    ],

    search: {
      provider: 'local'
    },

    sidebar: [
      {
        text: 'Introduction',
        items: [
          { text: 'Getting started', link: '/getting-started' },
        ],
        collapsed: false,
      },
      {
        text: 'Reference',
        items: [
          { text: 'Workbook', link: '/workbook' },
          { text: 'Worksheet', link: '/worksheet' },
        ],
        collapsed: false,
      },
      // {
      //   text: 'Examples',
      //   items: [
      //     { text: 'Examples', link: '/examples' },
      //   ],
      //   collapsed: false,
      // }
    ],

    socialLinks: [
      { icon: 'github', link: 'https://github.com/mikegaw/excelifyjs' },
      { icon: 'npm', link: 'https://www.npmjs.com/package/excelifyjs' }
    ],

    footer: {
      message: 'Released under the MIT License.',
      copyright: 'Copyright © 2026-present Michal Gawlowski'
    }
  }
})
