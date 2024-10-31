# 土壤分类样式生成工具

这是一个用于自动生成土壤分类系统样式方案的Python工具。该工具可以根据土壤分类的层级关系,自动为不同层级的土壤类型分配合适的样式,并保持样式的视觉层次感。

## 功能特点

- 支持多层级土壤分类(土类、亚类、土属、土种)的颜色分配
- 基于色系和色调的智能颜色插值
- 自动处理颜色冲突,确保不同分类使用不同颜色
- 支持RGB和HEX格式的颜色表示
- 结果可导出为Excel格式
- 并自动写入到stylex文件中

## 工作原理

1. 颜色分配策略:
   - 土类使用最深色
   - 亚类、土属、土种使用渐变色
   - 当需要的颜色数量超过现有色标时,自动进行颜色插值

2. 主要处理流程:
   - 创建颜色查找字典
   - 建立推荐颜色映射
   - 为每个层级分配颜色
   - 处理颜色冲突


## 注意事项

- 输入文件需符合指定格式
- 确保颜色推荐表中包含所有土类的推荐色系
- 建议在使用前检查输入数据的完整性
