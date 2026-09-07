# 快速启动

## 1. 运行参考内核

在系统根目录执行：

```bash
python runtime/xperience_runtime.py demo \
  --scenarios runtime/demo_scenarios.json \
  --out examples/demo_results.json
```

终端将返回场景数量、结果哈希和输出路径。

## 2. 运行测试

```bash
python -m unittest discover -s tests -v
```

预期：12 项测试通过（7 项运行时行为测试 + 5 项 JSON Schema 验证）。

## 3. 打开离线驾驶舱

双击 `prototype/index.html`。页面内嵌演示数据，无需网络、数据库和模型 API。

## 4. 阅读正式规范

`docs/段玉聪_DIKWP-Xperience_OS_人工主观体验生成与连续自我操作系统_v1.0_CN_最终版.docx`

## 5. 接入真实模型的最小路线

1. 选择开放权重多模态模型并暴露中间残差流或状态空间；
2. 接入 J-lens、SAE、线性探针或因果追踪作为工作区读写适配器；
3. 建立持续运行的人工身体状态总线；
4. 用小型循环状态模型实现可学习 Q-field；
5. 先做无报告、消融、价性反转和记忆断连四项预注册实验；
6. 在虚拟环境通过后再接入具身设备；
7. 外部行动始终通过 P-Space，体验本身不能自我授权。
