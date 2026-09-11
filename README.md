[README.md](https://github.com/user-attachments/files/32098464/README.md)
# DataLens Pro

一个完全在浏览器本地运行的数据分析与可视化工作台。

## 能做什么

- 导入 CSV、TSV、XLSX、XLS、JSON
- 自动识别字段类型并生成数据画像
- 表格搜索、筛选、排序和单元格编辑
- 数据清洗、计算列、撤销与恢复
- 图表配置、透视分析与报告导出
- 浅色 / 深色主题与响应式布局

## 快速开始

这是一个零构建依赖的静态前端项目。直接用浏览器打开 `index.html` 即可运行。

如果浏览器限制本地文件加载，可以启动任意静态服务器：

```bash
npx serve .
```

然后访问终端显示的本地地址。

## 项目结构

```text
.
├── index.html       # 页面结构与可访问性标记
├── src/
│   ├── app.js       # 数据状态、交互与分析逻辑
│   └── styles.css   # 设计系统、组件与响应式样式
├── package.json
└── .gitignore
```

## 技术说明

- 原生 HTML / CSS / JavaScript，无框架运行时
- SheetJS 用于表格文件解析与导出
- Apache ECharts 用于图表渲染
- 第三方库使用多 CDN 回退加载；粘贴文本、表格浏览和清洗在离线时仍可用
- 数据默认只留在当前浏览器，不上传到应用服务器

## License

MIT
