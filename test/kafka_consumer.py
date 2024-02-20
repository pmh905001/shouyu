# from kafka import KafkaConsumer

import kafka

consumer = kafka.KafkaConsumer('test2', bootstrap_servers=['localhost:9092'])

for message in consumer:
    print(message.value.decode('utf-8'))