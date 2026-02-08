import pickle
from datetime import datetime
from decimal import Decimal

# 创建包含datetime和Decimal对象的字典
data = {
    'date': datetime.now(),
    'amount': Decimal('123.45')
}

# 序列化数据
serialized_data = pickle.dumps(data)

# 反序列化数据
deserialized_data = pickle.loads(serialized_data)

# 验证类型
print(f"Type of deserialized 'date': {type(deserialized_data['date'])}")
print(f"Type of deserialized 'amount': {type(deserialized_data['amount'])}")