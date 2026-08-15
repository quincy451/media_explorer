CXX ?= g++
PKG_CONFIG ?= pkg-config

CPPFLAGS += -Isrc $(shell $(PKG_CONFIG) --cflags Qt5Widgets libvlc)
CXXFLAGS ?= -O2
CXXFLAGS += -std=c++17 -fPIC -Wall -Wextra -Wpedantic -MMD -MP
LDLIBS += $(shell $(PKG_CONFIG) --libs Qt5Widgets libvlc) -pthread

BUILD_DIR := build
DIST_DIR := dist/MediaExplorer
TARGET := $(DIST_DIR)/MediaExplorer
KEYBOARD_TEST := $(BUILD_DIR)/keyboard-acceptance
COMBINE_STATE_TEST := $(BUILD_DIR)/combine-job-state-test
PLAYBACK_SMOKE := $(BUILD_DIR)/playback-smoke

COMMON_SOURCES := \
	src/AppConfig.cpp \
	src/CombineJobState.cpp \
	src/MediaTools.cpp \
	src/MountDiscovery.cpp \
	src/MediaExplorerWindow.cpp
APP_SOURCES := $(COMMON_SOURCES) src/main.cpp
APP_OBJECTS := $(patsubst src/%.cpp,$(BUILD_DIR)/%.o,$(APP_SOURCES))
KEYBOARD_OBJECT := $(BUILD_DIR)/keyboard_acceptance.o
COMBINE_STATE_OBJECT := $(BUILD_DIR)/combine_job_state_test.o
PLAYBACK_SMOKE_OBJECT := $(BUILD_DIR)/playback_smoke.o
DEPENDENCIES := $(APP_OBJECTS:.o=.d) $(KEYBOARD_OBJECT:.o=.d) \
	$(COMBINE_STATE_OBJECT:.o=.d) $(PLAYBACK_SMOKE_OBJECT:.o=.d)

.PHONY: all clean test self-test keyboard-test combine-state-test playback-smoke install

all: $(TARGET)

$(TARGET): $(APP_OBJECTS)
	@mkdir -p $(dir $@)
	$(CXX) $(LDFLAGS) -o $@ $^ $(LDLIBS)

$(BUILD_DIR)/%.o: src/%.cpp
	@mkdir -p $(dir $@)
	$(CXX) $(CPPFLAGS) $(CXXFLAGS) -c -o $@ $<

$(KEYBOARD_OBJECT): tests/keyboard_acceptance.cpp
	@mkdir -p $(dir $@)
	$(CXX) $(CPPFLAGS) $(CXXFLAGS) -c -o $@ $<

$(COMBINE_STATE_OBJECT): tests/combine_job_state_test.cpp
	@mkdir -p $(dir $@)
	$(CXX) $(CPPFLAGS) $(CXXFLAGS) -c -o $@ $<

$(PLAYBACK_SMOKE_OBJECT): tests/playback_smoke.cpp
	@mkdir -p $(dir $@)
	$(CXX) $(CPPFLAGS) $(CXXFLAGS) -c -o $@ $<

$(KEYBOARD_TEST): $(KEYBOARD_OBJECT) $(filter-out $(BUILD_DIR)/main.o,$(APP_OBJECTS))
	$(CXX) $(LDFLAGS) -o $@ $^ $(LDLIBS)

$(COMBINE_STATE_TEST): $(COMBINE_STATE_OBJECT) $(BUILD_DIR)/CombineJobState.o
	$(CXX) $(LDFLAGS) -o $@ $^ $(LDLIBS)

$(PLAYBACK_SMOKE): $(PLAYBACK_SMOKE_OBJECT) $(filter-out $(BUILD_DIR)/main.o,$(APP_OBJECTS))
	$(CXX) $(LDFLAGS) -o $@ $^ $(LDLIBS)

self-test: $(TARGET)
	QT_QPA_PLATFORM=offscreen $(TARGET) --self-test

keyboard-test: $(KEYBOARD_TEST)
	QT_QPA_PLATFORM=offscreen $(KEYBOARD_TEST)

combine-state-test: $(COMBINE_STATE_TEST)
	$(COMBINE_STATE_TEST)

playback-smoke: $(PLAYBACK_SMOKE)
	@test -n "$(PLAYBACK_CLIP)" || { echo "Set PLAYBACK_CLIP=/path/to/video" >&2; exit 2; }
	QT_QPA_PLATFORM=xcb $(PLAYBACK_SMOKE) $(PLAYBACK_SMOKE_ARGS) "$(PLAYBACK_CLIP)"

test: self-test keyboard-test combine-state-test

install: $(TARGET)
	./install.sh

clean:
	rm -rf -- $(BUILD_DIR) $(DIST_DIR)

-include $(DEPENDENCIES)
