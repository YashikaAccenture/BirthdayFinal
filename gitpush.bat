
cd C:\Users\yashika.a.gupta\Desktop\WishingTool\Birthday

git add --all
git commit -m "wgitwatch autoCommit %date:~-4%%date:~3,2%%date:~0,2%.%time:~0,2%%time:~3,2%%time:~6,2%"
git pull https://github.com/YashikaAccenture/BirthdayFinal.git master
git push -f origin master
exit
