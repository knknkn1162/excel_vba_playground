.PHONY: commit push

ifeq (,$(XLSM))
	$(warning XLSM variable is not set!)
endif
push: commit
	git push
commit:
	git add $(XLSM)/
	git commit -m "$(COMMIT_MSG) $(XLSM)"
