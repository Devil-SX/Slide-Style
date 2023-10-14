# Slide Style

能够调整 PPTX 幻灯片图片的色调使其统一

# 原理

- 设置想要的主体色调
- 读取 PPTX 中的图片
- 去除黑色和白色
- 对其余像素点使用聚类算法
- 根据类别赋予颜色

# Get Started

- Pillow
- python-pptx
- scikit-learn
- numpy