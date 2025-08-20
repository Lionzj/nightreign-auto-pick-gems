## 这是一个通过AI生成的，黑夜君临自动挑选宝石遗物的小脚本

### 具体操作步骤：

#### 1.将宝石库里的宝石信息换成自己的。

​		注意只需要修改表里的颜色和词条列即可，序号按获取时间排序，获取时间越早的排越前。虽然前期录入现有宝石信息很费时间，但是全部录入完成后，宝石数量上来之后，基本也就每天新增个一两条数据的事情。录入前建议先将自己的宝石库筛选一下，把肯定不会用到的宝石先卖一卖以减少工作量。



#### 2.根据你想要的词条来修改对应职业脚本里的词条列表

注意词条内容需要跟前面Excel录入的完全一致。

比如:

```python
REQUIRED_PHRASES = {
        "每次打倒封印监牢里的囚犯，能永久性提升攻击力",
        "【铁之眼】技艺的使用次数+1",
        "【铁之眼】延长弱点暴露的时间",
        "加快累积绝招量表+3",
        "提升物理攻击力+2",
        "攻击成功时，能恢复精力",
        "出击时，会持有“石剑钥匙”"
}
```



#### 3.执行对应职业的py脚本

1. 将“宝石库.xlsx”放在脚本同目录下，确保包含“宝石库”工作表，且有“序号”、“颜色”、“词条一”、“词条二”、“词条三”列。
2. 运行脚本后，结果将保存至“铁之眼.xlsx”，并在控制台输出处理信息和组合数量。

```
lixiang@lixiangdeMacBook-Pro ~ % /usr/local/bin/python3 /Users/lixiang/Desktop/nightreign/铁之眼.py
开始处理...
load_and_preprocess executed in 0.08 seconds
find_valid_combinations executed in 6.94 seconds
export_to_excel executed in 0.01 seconds

结果已保存至: /Users/lixiang/Desktop/nightreign/铁之眼.xlsx
共找到 2 个符合条件的组合
```

