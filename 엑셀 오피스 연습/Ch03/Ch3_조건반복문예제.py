#
# i=50

# #if else 문
# if i > 20 :
#     print ("20보다 커요")
# elif i >10 :
#     print ("10보다 커요")
# else :
#     print ("10보다 작거나 같아요")

# #match case 문
# match i:
#     case num if 1<= num <10:
#         print ("한자리 수입니다")
#     case num if 10<= num <100:
#         print ("두자리 수입니다")
#     case _:
#         print ("세자리수 이상입니다")

# #while문
# i=1
# j=0
# while i<=100:
#     j=j+i
#     i=i+1
# print (j)

# #for 문
# j=0
# for i in range (100):
#     j=j+i
# print(j)

#에러 예외처리
i=10
j=1
try:
    k=i/j
    print (k)
except:
    print ("에러가 났어요")

print ("작업완료")
