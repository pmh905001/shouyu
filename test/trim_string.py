def truncate_string(s, length):
    # 将字符串转换为 Unicode
    # s = s.encode('utf-8')
    # 截取前 length 个字符
    truncated_s = s[:length]
    # 将截取后的字符串转换回字符串类型
    # truncated_s = truncated_s.decode('utf-8')
    return truncated_s

# 示例用法
s = "这是一个包含中文的字符串示例"
truncated_string = truncate_string(s, 100)
print(truncated_string)