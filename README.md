# 批量文件重命名工具使用说明

## 功能特性
- 支持正则表达式替换
- 提供多种命名规则模板
- 支持日期时间格式化
- 智能保留文件扩展名
- 支持多规则组合使用

## 命名规则说明

### 基础占位符
| 符号 | 说明                      |
|------|-------------------------|
| *    | 原文件名                 |
| #    | 自动序号（自动补零）      |
| $    | 随机数字                 |
| ?    | 原文件名的单个字符        |
| \n   | 原文件名第n个字符（从1开始）|

### 文件名截取
| 格式      | 示例             | 说明                    |
|-----------|------------------|-----------------------|
| `<*-n>`   | `<*-3>`          | 去掉末尾3个字符        |
| `<-n*>`   | `<-2*>`          | 去掉开头2个字符        |

### 规则字符转义
使用`\`转移直接输出规则符号：`\$` `\#` `\?`

## 日期时间格式化

### 标准格式代码
| 代码 | 说明                | 示例        |
|------|-------------------|------------|
| yyyy | 四位年份           | 2024       |
| yy   | 两位年份           | 24         |
| YYYY | 中文年份           | 二〇二四    |
| mm   | 补零月份           | 07         |
| m    | 不补零月份         | 7          |
| MM   | 中文月份           | 七月        |
| dd   | 补零日期           | 05         |
| d    | 不补零日期         | 5          |
| w/ddd    | 星期 (阿拉伯数字)      | 5          |
| W/DDD    | 星期 (中文数字)      | 五          |
| hh   | 24小时制           | 14         |
| h    | 12小时制           | 02 PM      |
| tt   | 补零分钟           | 08         |
| ss   | 秒钟               | 59         |

### 快捷日期格式
| 标签   | 格式示例                 |
|--------|-------------------------|
| `<->`  | 2025-3-19               |
| `<-->` | 2025年3月19日           |
| `<丨>`  | 二〇二五年三月十九日     |
| `<:>` | 0557（时分）           |
| `<::>` | 055712（时分秒）           |
| `<-:>` | 2025-3-19 0444           |
| `<.>`  | 20250319                |

## 使用示例

### 基础重命名
原文件名：`File_001.jpg`
- `<*-4>_新` → `File_新.jpg`
- `\1\2\3` → `Fil.jpg` 
- `Doc_#` → `Doc_01.jpg`

### 日期组合
原文件名：`photo.jpg`
- `拍摄于<YYYY年MM月DD日>` → `拍摄于二〇二四年七月零五日.jpg`
- `记录<-->< hh:tt>` → `记录2024年7月5日 14:08.jpg`

### 混合使用
原文件名：`DSC0001.jpg`
- `#_<mmdd>` → `01_0705.jpg`
- `$_<-->` → `5837_2024年7月5日.jpg`

---

# Batch File Renaming Tool User Guide

## Features
- Supports regular expression replacement
- Provides multiple naming rule templates
- Supports date and time formatting
- Automatically retains file extensions
- Supports the combination of multiple rules

## Naming Rule Explanation

### Basic Placeholders
| Symbol | Description                  |
|--------|-----------------------------|
| *      | Original filename            |
| #      | Auto sequence number (with leading zeros) |
| $      | Random number               |
| ?      | Single character from original filename |
| \n     | The nth character of the original filename (starting from 1) |

### Filename Trimming
| Format     | Example         | Description                |
|------------|-----------------|---------------------------|
| `<*-n>`    | `<*-3>`         | Remove the last 3 characters |
| `<-n*>`    | `<-2*>`         | Remove the first 2 characters |

### Escaping Rule Characters
Use `\` to escape and directly output rule symbols: `\$` `\#` `\?`

## Date and Time Formatting

### Standard Format Codes
| Code | Description              | Example        |
|------|-------------------------|----------------|
| yyyy | Four-digit year         | 2024           |
| yy   | Two-digit year          | 24             |
| YYYY | Chinese year            | 二〇二四        |
| mm   | Zero-padded month       | 07             |
| m    | Non-zero-padded month   | 7              |
| MM   | Chinese month           | 七月            |
| dd   | Zero-padded day         | 05             |
| d    | Non-zero-padded day     | 5              |
| w/ddd | Day of the week (Arabic numeral) | 5      |
| W/DDD | Day of the week (Chinese numeral) | 五      |
| hh   | 24-hour format          | 14             |
| h    | 12-hour format          | 02 PM          |
| tt   | Zero-padded minute      | 08             |
| ss   | Seconds                 | 59             |

### Quick Date Formats
| Tag   | Format Example                |
|-------|-------------------------------|
| `<->`  | 2025-3-19                     |
| `<-->` | 2025年3月19日                 |
| `<丨>` | 二〇二五年三月十九日           |
| `<:>`  | 0557 (hour-minute)            |
| `<::>` | 055712 (hour-minute-second)   |
| `<-:>` | 2025-3-19 0444                |
| `<.>`  | 20250319                      |

## Usage Examples

### Basic Renaming
Original filename: `File_001.jpg`
- `<*-4>_New` → `File_New.jpg`
- `\1\2\3` → `Fil.jpg` 
- `Doc_#` → `Doc_01.jpg`

### Date Combination
Original filename: `photo.jpg`
- `Taken on <YYYY-MM-DD>` → `Taken on 2024-07-05.jpg`
- `Recorded on <--> at <hh:tt>` → `Recorded on 2024-07-05 at 14:08.jpg`

### Mixed Usage
Original filename: `DSC0001.jpg`
- `#_<mmdd>` → `01_0705.jpg`
- `$_<-->` → `5837_2024年7月5日.jpg`

---

Developed by [yumumao@medu.cc](mailto:yumumao@medu.cc)
