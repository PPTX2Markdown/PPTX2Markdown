| Feature | pptx2markdown (rewrite) 0.2.0.dev0 | markitdown 0.1.9 | anydoc 0.2.5 |
| --- | :---: | :---: | :---: |
| Each slide is delimited | ✅ | ✅ | ❌ |
| Title placeholder becomes a heading | ✅ | ✅ | ✅ |
| Large first text box becomes a heading | ✅ | ❌ | ❌ |
| Two-column reading order | ✅ | ❌ | ❌ |
| Nested bullet levels | ✅ | ❌ | ✅ |
| Hyperlinks | ✅ | ❌ | ✅ |
| Bold runs | ✅ | ❌ | ✅ |
| Typed Markdown/HTML stays literal | ✅ | ❌ | ✅ |
| Tables (incl. a pipe in a cell) | ✅ | ❌ | ✅ |
| Merged cells keep row meaning | ✅ | ❌ | ❌ |
| Equations (OMML) as LaTeX | ✅ | ❌ | ✅ |
| Speaker notes | ✅ | ✅ | ✅ |
| Hidden slides are marked | ✅ | ❌ | ❌ |
| Image alt text | ✅ | ✅ | ✅ |
| Chart data | ✅ | ✅ | ✅ |
| Text inside grouped shapes | ✅ | ✅ | ✅ |
| Deck without slides | ✅ | ✅ | 💥 |
| Second column of a Two Content layout | ✅ | ✅ | ✅ |
| Header typed one character per line | ✅ | ❌ | ❌ |
| Strict Open XML deck | ✅ | 💥 | ✅ |
| **Passed** | **20/20** | **8/20** | **13/20** |

Time for a 300-slide deck (text, table per slide), Python 3.12.13, macOS-26.5.2:

| Converter | Seconds |
| --- | ---: |
| pptx2markdown (rewrite) | 0.224 |
| markitdown | 0.297 |
| anydoc | 0.012 |
