# 拓展：特殊排版（bignum、asym）

仅在用户选了"特殊排版"时使用。给渲染器副本追加两个 block，不修改 skill 原件里的 `scripts/render_pptx.py`。

## 接入步骤

1. 打开 `extensions/special-layout/blocks.py`，把 `_point_texts`、`_bignum`、`_asym` 三个函数粘进**渲染器副本**，位置在 `BLOCK_RENDERERS = {` 之前。文件头注释列出的依赖（颜色常量、`add_text`、`add_rect`、`add_hline`、`text_height` 等）副本里都已有，不要重复 import。
2. 在副本的 `BLOCK_RENDERERS` 里登记：`"bignum": _bignum, "asym": _asym,`。
3. 渲染器副本调用 `validate_deck` 时传入的是 `BLOCK_RENDERERS` 里的全部 kind，无需额外设置；单独跑 `deck_check.py` 时加参数：
   `python deck_check.py deck.json --extra-kinds bignum,asym`

## deck.json 字段

`bignum`：核心结论页，一个超大数字加右侧 2–3 条要点，整页只放它或再配一个 `insight`。

```json
{ "kind": "bignum", "kicker": "核心结论", "value": "37%", "title": "单次调用成本下降",
  "body": "对比统一计量前后三个月的账单，按同一口径折算。",
  "items": ["计量统一后重复采购的卡数减少", "配额回收让闲置的卡重新入池", "按调用量计费促使业务优化"] }
```

`asym`：7:3 非对称分栏，左主论点，右侧浅底放依据要点（`items` 为字符串，最多约 4 条）。

```json
{ "kind": "asym",
  "left": { "kicker": "主论点", "title": "先建计量，再谈治理", "body": "两三句说明……" },
  "right": { "kicker": "依据", "items": ["三个月账单对账", "九个业务的调用日志"] } }
```

## 使用原则

- 全稿 `bignum` 不超过 2 页，只用在真正的核心数字上；`value` 必须来自资料，不编造。
- `asym` 的 `left.body` 不要超过 4 行，依据要点每条一句。
- 这两个 kind 仍受相邻页版式不同的规则约束，不要连续放两页同类 block。
- `deck_check` 的字数统计会计入 `kicker`、`title`、`body`、`items`，不计 `value`。
