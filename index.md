---
layout: default
title: APDL 腳本目錄
---

# APDL 腳本目錄

這是一個 GitHub Pages 示範目錄。目前收錄三個可獨立執行的 Mechanical APDL 教學腳本。

> **執行前注意：** 三個示範腳本都以 `/CLEAR,NOSTART` 開始，執行時會清除目前的 MAPDL database。請先儲存正在工作的模型。

## 01｜BEAM188 懸臂梁

建立參數化矩形截面懸臂梁，在自由端施加集中力，求解後顯示變形並列出位移與反力。

- [在 GitHub 查看原始碼]({{ site.github.repository_url }}/blob/main/scripts/01_structural/beam188_cantilever.mac)
- [下載 `.mac` 腳本]({{ '/scripts/01_structural/beam188_cantilever.mac' | relative_url }})

## 02｜SHELL181 懸臂板

建立參數化薄板模型，固定一側並將總載重平均分配到另一側的節點。

- [在 GitHub 查看原始碼]({{ site.github.repository_url }}/blob/main/scripts/02_structural/shell181_cantilever_plate.mac)
- [下載 `.mac` 腳本]({{ '/scripts/02_structural/shell181_cantilever_plate.mac' | relative_url }})

## 03｜SOLID278 穩態熱傳

建立三維長方體熱傳模型，在兩端指定溫度，計算穩態溫度分布。

- [在 GitHub 查看原始碼]({{ site.github.repository_url }}/blob/main/scripts/03_thermal/solid278_steady_conduction.mac)
- [下載 `.mac` 腳本]({{ '/scripts/03_thermal/solid278_steady_conduction.mac' | relative_url }})

## 使用方式

1. 下載需要的 `.mac` 檔案。
2. 在 Mechanical APDL 中選擇 `Utility Menu > File > Read Input from...`。
3. 選取腳本並執行。
4. 先修改腳本開頭的參數區，即可測試不同尺寸、材料與載重。

這些檔案只用於語法與網站目錄示範，不代表正式工程分析設定。

