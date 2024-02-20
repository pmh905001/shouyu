# from kafka import KafkaProducer

import kafka

producer = kafka.KafkaProducer(bootstrap_servers=['localhost:9092'])

message = 'This is a message from Python2'
producer.send('test2', message.encode('utf-8'))

producer.close()