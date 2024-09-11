import yaml

print("hello world")

with open("excel/info.yaml", "r") as f:
    data = yaml.load(f, Loader=yaml.SafeLoader)

print(data)