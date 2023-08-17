from queue import Queue


class TaskExecutor:

    def __init__(self):
        self.queue = Queue(20)

    def add(self, fun, args):
        self.queue.put((fun, args))

    def run(self):
        while True:
            msg = self.queue.get()
            fun = msg[0]
            fun(*msg[1])
            # time.sleep(1)
            self.queue.task_done()
