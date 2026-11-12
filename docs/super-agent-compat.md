# dsh-super-agent 0.0.5 隔离兼容性试验

日期：2026-11-11
目的：在不影响 `desktop` profile 的前提下，验证 `dsh-super-agent@0.0.5` 能否跑在
DSH 0.2.2-rc.1 上。上一次直接装进 desktop 曾导致无法启动，这次改在独立 profile 里做。

## 试验步骤

```powershell
$dsh = 'D:\DSH\resources\runtime\cli\bin\dsh.cmd'
# 1. 从官方 web 模板生成独立 profile（不动 desktop）
& $dsh --profile superagent --from-default-profile web --dump-config
# 2. 只往这个 profile 装
& $dsh plugin --profile superagent add dsh-super-agent@0.0.5
# 3. 真正启动
& $dsh --profile superagent
```

`dsh plugin add` 会自动把 `dsh-super-agent` 注册进该 profile 的
`dsh.profile.bundles`，装完的依赖树里被提升到 profile 顶层的是**旧核心包**：

```
@deepseek-ai/dsh-session            0.1.2-rc.1
@deepseek-ai/dsh-llm                0.1.2-rc.1
@deepseek-ai/dsh-agent-presets      0.1.2-rc.1
@deepseek-ai/dsh-typert-protocol    0.1.2-rc.1
@deepseek-ai/schemastery            3.19.0
```

## 结果：启动失败，原因清楚

启动时 DSH 主动拒绝了两个核心行：

```
dsh: disabling profile plugin row "session": Plugin @deepseek-ai/dsh-session@0.1.2-rc.1
  is incompatible with dsh 0.2.2-rc.1: requirements {"@deepseek-ai/dsh-scope":"^0.1.2-rc.1"}
dsh: disabling profile plugin row "session-projection": Plugin @deepseek-ai/dsh-session-projection@0.1.2-rc.1
  is incompatible with dsh 0.2.2-rc.1: requirements {"@deepseek-ai/dsh-session":"^0.1.2-rc.1"}
```

随后 boot 直接失败：

```
dsh: startup failed: 1 required plugin did not activate
Failed plugins (1):
  typert-loader
    AggregateError: typert-loader: 1 typert contributor(s) failed to register:
      - @deepseek-ai/dsh-llm invocation "@deepseek-ai/dsh-llm#llm/discoverModels"
        parameter codec has no create() factory
```

另外 `--dump-config` 阶段还有一条 patch 警告：`patch: entry "agent-presets" not found`
——它的 `cordis.patch.yml` 假设 web 模板里存在 `agent-presets` 行，但 0.2.2-rc.1 的
组合里行 id 不同，这段 preset 注入没有生效。

完整日志：`docs/super-agent-boot.log`。

## 根因

`dsh-super-agent@0.0.5` 的依赖锁定在 `@deepseek-ai/dsh-* ^0.1.2-rc.1`，装进 profile 后
这些旧包被提升到 profile 顶层，profile 内的解析优先命中 0.1.2-rc.1，而运行时是
0.2.2-rc.1。两者在 typert 编解码契约上已不兼容（新版要求 codec 提供 `create()` 工厂），
所以 `typert-loader` 注册失败、启动中止。

这不是"少一个配置"能补的：要修必须由上游把依赖升级到 0.2.x 并重新发布。

## 结论与建议

- **0.0.5 无法在 0.2.2-rc.1 上运行**；装进常驻 desktop profile 不只是"可能坏"，
  而是会把 `session` 等核心行拖下水，直接导致应用无法启动（已发生过一次）。
- 隔离 profile 的试验是干净的：失败被限制在 `superagent` 里，desktop 全程正常。
- 试验后已删除 `C:\Users\Administrator\.dsh\profiles\superagent`，桌面 profile 恢复为
  `dsh-base + dsh-web-app + dsh-filens`。
- 等上游发布适配 0.2.x 的版本后，重跑本文的三条命令即可复验。
