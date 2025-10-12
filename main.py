from collections import deque
import pygame
import sys
import time
import os
import math
import random
import heapq
from openpyxl import Workbook, load_workbook

# Initialize Pygame
pygame.init()
try:
    pygame.mixer.init()
    if os.path.exists("game_theme.mp3"):
        pygame.mixer.music.load("game_theme.mp3")
        pygame.mixer.music.play(-1)  # loop
    else:
        print("Audio disabled: game_theme.mp3 not found")
except Exception as e:
    print("Audio disabled:", e)
# ---------------------------
# Constants / Colors
# ---------------------------
# Use larger grid so levels 4-5 fit; earlier levels just draw inside this area.
GRID_WIDTH = 16
GRID_HEIGHT = 10
CELL_SIZE = 60
WINDOW_WIDTH = GRID_WIDTH * CELL_SIZE + 350
WINDOW_HEIGHT = GRID_HEIGHT * CELL_SIZE + 100
FPS = 60

SPACE_BLACK = (10, 10, 20)
ICE_BLUE = (173, 216, 255)
WALL_GRAY = (60, 70, 85)
DARK_GRAY = (30, 35, 45)
REBEL_ORANGE = (255, 165, 0)
EMPIRE_RED = (220, 20, 60)
JEDI_GREEN = (50, 205, 50)
SABER_BLUE = (64, 164, 255)
FORCE_PURPLE = (138, 43, 226)
HOTH_WHITE = (245, 245, 255)
CONSOLE_GREEN = (0, 255, 127)
HOLOGRAM_CYAN = (0, 255, 255)
KEY_GOLD = (255, 215, 0)
A_STAR_YELLOW = (255, 165, 0)     # path/expansion for A*
GREEDY_PINK   = (255, 105, 180)     # path/expansion for Greedy (hot pink)

# ---------------------------
# Maze (MULTI-LEVEL)
# ---------------------------
class Maze:
    """Maze with 5 levels and a collectible key each level"""
    def __init__(self, level=1):
        self.level = level
        self.key_collected = False
        self.key_pos = None
        self.enemy_spawns = []
        self.load_level(level)

    def load_level(self, level):
        self.level = level
        self.key_collected = False
        self.enemy_spawns = []

        if level == 1:
            # 12 x 9
            self.grid = [
                [1,1,1,1,1,1,1,1,1,1,1,1],
                [1,0,0,0,1,0,0,0,0,0,0,1],
                [1,0,1,0,1,0,1,1,1,0,0,1],
                [1,0,0,0,0,0,0,0,1,0,0,1],
                [1,1,1,0,1,1,0,0,0,0,1,1],
                [1,0,0,0,0,0,0,1,0,0,0,1],
                [1,0,1,1,0,1,0,1,0,1,0,1],
                [1,0,0,0,0,0,0,0,0,0,0,1],
                [1,1,1,1,1,1,1,1,1,1,1,1]
            ]
            self.start_pos = (1, 1)
            self.goal_pos  = (10, 7)
            self.key_pos   = (5, 3)
            self.enemy_spawns = []  # no enemies on level 1

        elif level == 2:
            # 12 x 9 — harder than 1; side trip to key
            self.grid = [
                [1,1,1,1,1,1,1,1,1,1,1,1],
                [1,0,0,0,1,0,0,0,1,0,0,1],
                [1,0,1,0,1,0,1,0,1,0,0,1],
                [1,0,1,0,0,0,1,0,0,1,0,1],
                [1,0,1,1,1,0,1,1,0,1,0,1],
                [1,0,0,0,0,0,0,1,0,0,0,1],
                [1,1,0,1,1,1,0,1,1,1,0,1],
                [1,0,0,0,0,0,0,0,0,0,0,1],
                [1,1,1,1,1,1,1,1,1,1,1,1]
            ]
            self.start_pos = (1, 1)
            self.goal_pos  = (10, 7)
            self.key_pos   = (9, 2)
            # (x, y, dx, dy) patrols (move 1 tile, bounce)
            self.enemy_spawns = [
                (2, 7, 1, 0),   # bottom corridor
                (5, 1, 0, 1),   # mid-right vertical shaft
            ]

        elif level == 3:
            # 12 x 9 — hardest of small boards
            self.grid = [
                [1,1,1,1,1,1,1,1,1,1,1,1],
                [1,0,0,1,0,0,0,1,1,0,0,1],
                [1,0,1,0,0,1,0,1,1,0,1,1],
                [1,0,1,0,0,0,0,0,1,0,0,1],
                [1,0,0,0,1,1,1,0,1,1,0,1],
                [1,1,1,0,0,0,1,0,0,1,0,1],
                [1,0,0,0,1,0,1,1,0,1,0,1],
                [1,0,0,0,0,0,0,0,0,0,0,1],
                [1,1,1,0,1,1,1,1,1,1,1,1]
            ]
            self.start_pos = (1, 7)
            self.goal_pos  = (10, 1)
            self.key_pos   = (4, 1)
            self.enemy_spawns = [
                (3, 3, 1, 0),
                (9, 6, 0, -1),
                (6, 5, -1, 0),
            ]

        elif level == 4:
            # 16 x 10 — bigger & tougher, still fair
            self.grid = [
                [1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1],
                [1,0,0,0,0,1,0,0,0,0,1,0,0,0,0,1],
                [1,0,1,1,0,1,0,1,1,0,1,0,1,1,0,1],
                [1,0,0,1,0,0,0,1,0,0,0,0,1,0,0,1],
                [1,1,0,1,1,1,0,1,0,1,1,0,1,0,1,1],
                [1,0,0,0,0,0,0,1,0,0,0,0,1,0,0,1],
                [1,0,1,1,0,1,0,1,1,1,0,1,1,1,0,1],
                [1,0,0,0,0,0,0,0,0,1,0,0,0,0,0,1],
                [1,0,1,1,1,1,1,1,0,1,1,1,1,1,0,1],
                [1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1],
            ]
            self.start_pos = (1, 1)
            self.goal_pos  = (14, 8)
            self.key_pos   = (1, 5)
            self.enemy_spawns = [
                (3, 7, 1, 0),   # long bottom-ish lane
                (12, 2, 0, 1),  # right column up↕down
                (7, 5, -1, 0),  # middle lane
            ]

        else:  # level 5 — biggest and hardest
            # 16 x 10
            self.grid = [
                [1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1],
                [1,0,0,0,1,0,0,0,0,1,0,0,0,0,0,1],
                [1,0,1,0,1,0,1,1,0,1,0,1,1,1,0,1],
                [1,0,1,0,0,0,0,1,0,0,0,0,1,0,0,1],
                [1,0,1,1,1,1,0,1,1,1,1,0,1,0,1,1],
                [1,0,0,0,0,0,0,0,0,0,0,0,1,0,0,1],
                [1,1,1,0,1,1,1,1,0,1,1,0,1,1,0,1],
                [1,0,0,0,0,0,0,0,0,1,0,0,0,0,0,1],
                [1,0,1,1,1,1,1,1,0,1,1,1,1,1,0,1],
                [1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1],
            ]
            self.start_pos = (1, 8)
            self.goal_pos  = (14, 1)
            self.key_pos   = (6, 3)
            self.enemy_spawns = [
                (4, 5, 1, 0),
                (10, 3, 0, 1),
                (13, 7, -1, 0),
                (8, 1, 0, 1),
            ]

        self.terrain_type = "ice"

    def is_wall(self, x, y):
        if 0 <= y < len(self.grid) and 0 <= x < len(self.grid[0]):
            return self.grid[y][x] == 1
        return True

    def slide_move(self, start_x, start_y, direction):
        x, y = start_x, start_y
        direction_map = {'up':(0,-1),'down':(0,1),'left':(-1,0),'right':(1,0)}
        if direction not in direction_map:
            return None
        dx, dy = direction_map[direction]
        moves = 0
        while True:
            nx, ny = x + dx, y + dy
            if self.is_wall(nx, ny): break
            x, y = nx, ny
            moves += 1
            if moves > 50: break
        return (x, y) if (x, y) != (start_x, start_y) else None

    def collect_key(self, player_pos):
        if not self.key_collected and player_pos == self.key_pos:
            self.key_collected = True
            return True
        return False

    def can_exit(self, player_pos):
        return self.key_collected and player_pos == self.goal_pos

    def get_level_name(self):
        return {
            1:"Hoth Landing Site",
            2:"Ice Caverns",
            3:"Echo Base Approach",
            4:"Frozen Wastes",
            5:"Shield Generator Run",
        }[self.level]

# ---------------------------
# Blind search (BFS/DFS)
# ---------------------------
class SearchAlgorithm:
    def __init__(self, maze):
        self.maze = maze
        self.explored = set()        # for UI coloring only (positions)
        self.path = []               # final path (list of positions)
        self.search_order = []       # order positions were expanded (for viz)
        self.algorithm_used = None
        self.nodes_expanded = 0

    def reset(self):
        self.explored.clear()
        self.path.clear()
        self.search_order.clear()
        self.algorithm_used = None
        self.nodes_expanded = 0

    def get_neighbors(self, pos):
        x, y = pos
        neighbors = []
        for d in ('up', 'down', 'left', 'right'):
            np = self.maze.slide_move(x, y, d)
            if np:
                neighbors.append(np)
        return neighbors

    def bfs_with_key(self, start, goal, key_pos, has_key_start=False):
        self.reset()
        self.algorithm_used = "BFS with Key Collection"

        q = deque([(start, has_key_start, [start])])
        seen = {(start, has_key_start)}

        while q:
            cur, has_key, path = q.popleft()
            self.search_order.append(cur)
            self.nodes_expanded += 1
            self.explored.add(cur)

            if cur == key_pos:
                has_key = True
            if cur == goal and has_key:
                self.path = path
                return True

            for nb in self.get_neighbors(cur):
                if nb == goal and not has_key:
                    continue
                st = (nb, has_key)
                if st not in seen:
                    seen.add(st)
                    q.append((nb, has_key, path + [nb]))
        return False

    def dfs_with_key(self, start, goal, key_pos, has_key_start=False):
        self.reset()
        self.algorithm_used = "DFS with Key Collection"

        stack = [(start, has_key_start, [start])]
        seen = set()

        while stack:
            cur, has_key, path = stack.pop()
            st = (cur, has_key)
            if st in seen:
                continue
            seen.add(st)

            self.search_order.append(cur)
            self.nodes_expanded += 1
            self.explored.add(cur)

            if cur == key_pos:
                has_key = True
            if cur == goal and has_key:
                self.path = path
                return True

            for nb in reversed(self.get_neighbors(cur)):
                if nb == goal and not has_key:
                    continue
                nst = (nb, has_key)
                if nst not in seen:
                    stack.append((nb, has_key, path + [nb]))
        return False
    
    def heuristic(self, a, b, heuristic_type="manhattan"):
        if heuristic_type == "euclidean":
            return math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2)
        else:  # default manhattan
            return abs(a[0] - b[0]) + abs(a[1] - b[1])

    def a_star_with_key(self, start, goal, key_pos, has_key_start=False, heuristic_type="manhattan"):
        self.reset()
        self.algorithm_used = f"A* with Key ({heuristic_type.title()})"

        pq = []
        heapq.heappush(pq, (0, start, has_key_start, [start]))
        g_cost = {(start, has_key_start): 0}
        seen = set()

        while pq:
            _, cur, has_key, path = heapq.heappop(pq)
            st = (cur, has_key)
            if st in seen:
                continue
            seen.add(st)

            self.search_order.append(cur)
            self.nodes_expanded += 1
            self.explored.add(cur)

            # When key is picked up, ensure new state exists in g_cost
            if cur == key_pos and not has_key:
                has_key = True
                g_cost[(cur, has_key)] = g_cost.get((cur, False), 0)

            if cur == goal and has_key:
                self.path = path
                return True

            for nb in self.get_neighbors(cur):
                if nb == goal and not has_key:
                    continue
                new_cost = g_cost[(cur, has_key)] + 1
                nst = (nb, has_key)
                if nst not in g_cost or new_cost < g_cost[nst]:
                    g_cost[nst] = new_cost
                    priority = new_cost + self.heuristic(nb, goal, heuristic_type)
                    heapq.heappush(pq, (priority, nb, has_key, path + [nb]))
        return False

    def greedy_with_key(self, start, goal, key_pos, has_key_start=False, heuristic_type="manhattan"):
        self.reset()
        self.algorithm_used = f"Greedy with Key ({heuristic_type.title()})"

        pq = []
        heapq.heappush(pq, (0, start, has_key_start, [start]))
        seen = set()

        while pq:
            _, cur, has_key, path = heapq.heappop(pq)
            st = (cur, has_key)
            if st in seen:
                continue
            seen.add(st)

            self.search_order.append(cur)
            self.nodes_expanded += 1
            self.explored.add(cur)

            if cur == key_pos:
                has_key = True
            if cur == goal and has_key:
                self.path = path
                return True

            for nb in self.get_neighbors(cur):
                if nb == goal and not has_key:
                    continue
                nst = (nb, has_key)
                if nst not in seen:
                    priority = self.heuristic(nb, goal, heuristic_type)
                    heapq.heappush(pq, (priority, nb, has_key, path + [nb]))
        return False

# ---------------------------
# Game
# ---------------------------
class StarWarsIceMazeGame:
    def __init__(self):
        self.screen = pygame.display.set_mode((WINDOW_WIDTH, WINDOW_HEIGHT))
        pygame.display.set_caption("Star Wars: Hoth Ice Maze - Multi-Level Mission")
        self.clock = pygame.time.Clock()

        # Fonts
        self.font = pygame.font.Font(None, 24)
        self.title_font = pygame.font.Font(None, 32)
        self.console_font = pygame.font.Font(None, 20)

        # Level state
        self.current_level = 1
        self.max_level = 5

        # Images
        self.images = self.load_images()

        # World + search
        self.maze = Maze(self.current_level)
        self.search = SearchAlgorithm(self.maze)

        # Game state
        self.player_pos = self.maze.start_pos
        self.manual_mode = True
        self.current_algorithm = None
        self.heuristic_used = 'manhattan'
        self.visualization_step = 0
        self.animating = False
        self.animation_speed = 8
        self.last_step_time = 0
        self.solution_found = False
        self.drawn_path = False
        self.game_completed = False
        self.show_start_screen = True

        # Enemies
        self.enemies = []                 # dicts: {'x','y','dx','dy'}
        self.enemy_step_interval = 0.35
        self._enemy_last_step = 0.0
        self.game_over = False
        self._load_enemies()

        # Autopilot (optional after AI finds path)
        self.autopilot = False
        self.autopath = []
        self.auto_index = 0
        self.autopilot_speed = 0.15
        self._auto_last_step = 0.0

        # Visual effects
        self.glow_effect = 0
        self.scan_lines = 0
        self.key_glow = 0

    # ---------- assets ----------
    def load_images(self):
        images = {}
        for name, filename in {'player':'pic1.png','base':'pic2.png','key':'key.png'}.items():
            try:
                if os.path.exists(filename):
                    img = pygame.image.load(filename).convert_alpha()
                    images[name] = pygame.transform.scale(img, (CELL_SIZE-10, CELL_SIZE-10))
                else:
                    images[name] = None
            except Exception:
                images[name] = None
        return images
    
    # ---------- Comparisons ------------

    def benchmark_algorithms(self, filename="search_benchmark.xlsx"):
        # Algorithms to run (Greedy Euclidean re-added)
        start_pos = self.maze.start_pos
        goal_pos  = self.maze.goal_pos
        key_pos   = self.maze.key_pos
        has_key   = False
        algorithms = [
            ("BFS", self.search.bfs_with_key, {
                "start": start_pos, "goal": goal_pos,
                "key_pos": key_pos, "has_key_start": has_key}),
            ("DFS", self.search.dfs_with_key, {
                "start": start_pos, "goal": goal_pos,
                "key_pos": key_pos, "has_key_start": has_key}),
            ("A* (Manhattan)", self.search.a_star_with_key, {
                "start": start_pos, "goal": goal_pos,
                "key_pos": key_pos, "has_key_start": has_key,
                "heuristic_type": "manhattan"}),
            ("A* (Euclidean)", self.search.a_star_with_key, {
                "start": start_pos, "goal": goal_pos,
                "key_pos": key_pos, "has_key_start": has_key,
                "heuristic_type": "euclidean"}),
            ("Greedy (Manhattan)", self.search.greedy_with_key, {
                "start": start_pos, "goal": goal_pos,
                "key_pos": key_pos, "has_key_start": has_key,
                "heuristic_type": "manhattan"}),
            ("Greedy (Euclidean)", self.search.greedy_with_key, {
                "start": start_pos, "goal": goal_pos,
                "key_pos": key_pos, "has_key_start": has_key,
                "heuristic_type": "euclidean"}),
        ]

        results = []

        print(f"\n===== Running Benchmarks for Level {self.current_level} =====\n")

        for name, func, kwargs in algorithms:
            # Run algorithm
            start_t = time.perf_counter()
            func(**kwargs)  # search functions update self.search.path and self.search.nodes_expanded
            runtime = round(time.perf_counter() - start_t, 6)

            # Read metrics from the search object (safer than assuming a return value)
            nodes_expanded = getattr(self.search, "nodes_expanded", None)
            path_length = len(getattr(self.search, "path", [])) if getattr(self.search, "path", None) is not None else 0

            results.append((self.current_level, name, runtime, nodes_expanded, path_length))
            print(f"{name:<25} | Runtime: {runtime:.6f}s | Nodes Expanded: {nodes_expanded} | Path Len: {path_length}")

        # Append results to Excel file (no timestamps)
        file_exists = os.path.exists(filename)
        if file_exists:
            wb = load_workbook(filename)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(["Level", "Algorithm", "Runtime (s)", "Nodes Expanded", "Path Length"])

        for level, name, runtime, nodes, path_len in results:
            ws.append([level, name, runtime, nodes, path_len])

        wb.save(filename)
        print(f"\n✅ Results saved to: {filename}\n")

    # ---------- Enemies ----------
    def _load_enemies(self):
        self.enemies = []
        for sx, sy, dx, dy in getattr(self.maze, "enemy_spawns", []):
            self.enemies.append({"x": sx, "y": sy, "dx": dx, "dy": dy})
        self._enemy_last_step = 0.0
        self.game_over = False

    def _step_enemies(self):
        # Enemies always move; collisions only matter in HUMAN mode.
        if not self.enemies or self.game_completed or self.game_over:
            return
        now = time.time()
        if now - self._enemy_last_step < self.enemy_step_interval:
            return
        self._enemy_last_step = now

        for e in self.enemies:
            nx, ny = e["x"] + e["dx"], e["y"] + e["dy"]

            if self.maze.is_wall(nx, ny):
                e["dx"] *= -1
                e["dy"] *= -1
                nx, ny = e["x"] + e["dx"], e["y"] + e["dy"]
                if self.maze.is_wall(nx, ny):
                    for dx, dy in [(1,0), (-1,0), (0,1), (0,-1)]:
                        tx, ty = e["x"] + dx, e["y"] + dy
                        if not self.maze.is_wall(tx, ty):
                            e["dx"], e["dy"] = dx, dy
                            nx, ny = tx, ty
                            break
                    else:
                        continue  # stuck

            e["x"], e["y"] = nx, ny

            # collision with player only in HUMAN mode
            if self.manual_mode and (e["x"], e["y"]) == self.player_pos:
                self._trigger_game_over()

    def _trigger_game_over(self):
        self.game_over = True
        self.animating = False
        self.autopilot = False
        print("❌ Game Over: a patrol caught you!")

    # ---------- Level helpers ----------
    def load_new_level(self):
        self.maze = Maze(self.current_level)
        self.search = SearchAlgorithm(self.maze)
        self.player_pos = self.maze.start_pos
        self.reset_search_ui_flags()
        self._load_enemies()

    def complete_level(self):
        if self.current_level < self.max_level:
            print(f"Level {self.current_level} complete! Moving to {self.current_level+1}.")
            self.current_level += 1
            self.load_new_level()
        else:
            self.game_completed = True
            print("All levels completed! You've reached Echo Base!")

    def start_new_game(self):
        self.current_level = 1
        self.load_new_level()
        self.show_start_screen = True
        self.game_completed = False
        print("Starting new game at Level 1...")

    def reset_search_ui_flags(self):
        self.search.reset()
        self.current_algorithm = None
        self.visualization_step = 0
        self.animating = False
        self.solution_found = False
        self.drawn_path = False
        self.game_over = False
        self.autopilot = False
        self.autopath = []
        self.auto_index = 0

    # ---------- Input ----------
    def handle_events(self):
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                return False

            if event.type == pygame.KEYDOWN:

                # --- Start screen ---
                if self.show_start_screen:
                    if event.key == pygame.K_1:
                        self.manual_mode = True
                        self.show_start_screen = False
                    elif event.key == pygame.K_2:
                        self.manual_mode = False
                        self.show_start_screen = False
                    elif event.key == pygame.K_ESCAPE:
                        return False
                    return True

                # --- Game over screen ---
                if self.game_over:
                    if event.key in (pygame.K_r, pygame.K_SPACE):
                        self.load_new_level()  # restart same level (mode preserved)
                    elif event.key == pygame.K_ESCAPE:
                        self.start_new_game()
                    return True

                # --- End screen ---
                if self.game_completed:
                    if event.key in (pygame.K_r, pygame.K_SPACE):
                        self.start_new_game()
                    elif event.key == pygame.K_ESCAPE:
                        self.start_new_game()
                    return True

                # --- In-game controls (Human) ---
                if self.manual_mode and not self.animating and not self.autopilot:
                    new_pos = None
                    if event.key in (pygame.K_UP, pygame.K_w):
                        new_pos = self.maze.slide_move(*self.player_pos, 'up')
                    elif event.key in (pygame.K_DOWN, pygame.K_s):
                        new_pos = self.maze.slide_move(*self.player_pos, 'down')
                    elif event.key in (pygame.K_LEFT, pygame.K_a):
                        new_pos = self.maze.slide_move(*self.player_pos, 'left')
                    elif event.key in (pygame.K_RIGHT, pygame.K_d):
                        new_pos = self.maze.slide_move(*self.player_pos, 'right')

                    if new_pos:
                        self.player_pos = new_pos

                        # collide with enemy after moving (human only)
                        for e in self.enemies:
                            if (e["x"], e["y"]) == self.player_pos:
                                self._trigger_game_over()
                                break
                        if self.game_over:
                            return True

                        if self.maze.collect_key(self.player_pos):
                            print("Key collected!")
                        if self.maze.can_exit(self.player_pos):
                            self.complete_level()

                # --- Algorithm controls (AI mode only) ---
                if not self.manual_mode:
                    if event.key == pygame.K_b and not self.animating:
                        self.start_bfs()
                    elif event.key == pygame.K_d and not self.animating:
                        self.start_dfs()
                    elif event.key == pygame.K_u and not self.animating:
                        self.start_a_star(heuristic_type=self.heuristic_used)
                    elif event.key == pygame.K_i and not self.animating:
                        self.start_greedy(heuristic_type=self.heuristic_used)
                    elif event.key == pygame.K_o and not self.animating:
                        if self.heuristic_used == 'manhattan':
                            self.heuristic_used = 'euclidean'
                        else:
                            self.heuristic_used = 'manhattan'
                        print("Changed to "+self.heuristic_used)
                    elif event.key == pygame.K_p and not self.animating:
                        self.benchmark_algorithms()
                    elif event.key == pygame.K_a and self.solution_found and not self.autopilot and self.drawn_path:
                        if self.search.path:
                            self.autopilot = True
                            self.autopath = list(self.search.path)
                            self.auto_index = 0
                            self.player_pos = self.autopath[0]
                            print("Autopilot engaged.")

                # Works in both modes
                if event.key == pygame.K_r:
                    self.load_new_level()  # mode preserved; search state reset
                elif event.key == pygame.K_m:
                    # toggle Human/AI (cancels autopilot/scan animation)
                    self.manual_mode = not self.manual_mode
                    self.animating = False
                    self.autopilot = False
                    if self.manual_mode:
                        # clear search viz if returning to Human
                        self.search.reset()
                        self.solution_found = False
                elif event.key == pygame.K_ESCAPE:
                    self.start_new_game()

        return True

    # ---------- Search triggers ----------
    def start_bfs(self):
        self.current_algorithm = "REBEL SCANNER (BFS)"
        self.solution_found = self.search.bfs_with_key(
            start=self.player_pos,
            goal=self.maze.goal_pos,
            key_pos=self.maze.key_pos,
            has_key_start=self.maze.key_collected
        )
        self.visualization_step = 0
        self.animating = True
        self.manual_mode = False

    def start_dfs(self):
        self.current_algorithm = "EMPIRE PROBE (DFS)"
        self.solution_found = self.search.dfs_with_key(
            start=self.player_pos,
            goal=self.maze.goal_pos,
            key_pos=self.maze.key_pos,
            has_key_start=self.maze.key_collected
        )
        self.visualization_step = 0
        self.animating = True
        self.manual_mode = False

    def start_a_star(self, heuristic_type="manhattan"):
        self.current_algorithm = f"RESISTANCE PATHFINDER (A*) [{heuristic_type.title()}]"
        self.solution_found = self.search.a_star_with_key(
            start=self.player_pos,
            goal=self.maze.goal_pos,
            key_pos=self.maze.key_pos,
            has_key_start=self.maze.key_collected,
            heuristic_type=heuristic_type
        )
        self.visualization_step = 0
        self.animating = True
        self.manual_mode = False

    def start_greedy(self, heuristic_type="manhattan"):
        self.current_algorithm = f"IMPERIAL SEEKER (Greedy) [{heuristic_type.title()}]"
        self.solution_found = self.search.greedy_with_key(
            start=self.player_pos,
            goal=self.maze.goal_pos,
            key_pos=self.maze.key_pos,
            has_key_start=self.maze.key_collected,
            heuristic_type=heuristic_type
        )
        self.visualization_step = 0
        self.animating = True
        self.manual_mode = False

    # ---------- Update / Draw ----------
    def update(self):
        current_time = time.time()
        self.glow_effect = (self.glow_effect + 3) % 360
        self.scan_lines = (self.scan_lines + 2) % WINDOW_HEIGHT
        self.key_glow = (self.key_glow + 5) % 360

        # animate search expansion
        if self.animating and current_time - self.last_step_time > (1.0 / self.animation_speed):
            if self.visualization_step < len(self.search.search_order):
                self.visualization_step += 1
                self.last_step_time = current_time
            else:
                self.animating = False  # finished scanning

        # Autopilot stepping
        player_blocked = False

        if self.autopilot and (current_time - self._auto_last_step) > self.autopilot_speed:
            self._auto_last_step = current_time
            if self.auto_index + 1 < len(self.autopath):
                # check if enemy is currently there OR could move there next step
                for e in self.enemies:
                    # predict where enemy will move
                    next_x = e["x"] + e["dx"]
                    next_y = e["y"] + e["dy"]

                    # if that tile is a wall, enemy will bounce back
                    if self.maze.is_wall(next_x, next_y):
                        next_x = e["x"] - e["dx"]
                        next_y = e["y"] - e["dy"]

                    # block if enemy is currently there OR will move there next step
                    if (e["x"], e["y"]) == self.autopath[self.auto_index+1] \
                    or (next_x, next_y) == self.autopath[self.auto_index+1]:
                        print("Enemy will block the next tile")
                        player_blocked = True
                        break



                if not player_blocked:
                    self.auto_index += 1
                    self.player_pos = self.autopath[self.auto_index]
                    self.maze.collect_key(self.player_pos)
                    if self.maze.can_exit(self.player_pos):
                        self.autopilot = False
                        self.complete_level()
            else:
                self.autopilot = False

        # step enemies (they move in all modes; only Human collides)
        if not self.animating:
            self._step_enemies()

        # ✅ universal collision check (after player + enemies)
        if not player_blocked and self.autopilot:  # player didn’t move, but still needs collision check
            for e in self.enemies:
                if (e["x"], e["y"]) == self.player_pos:
                    self._trigger_game_over()
                    break

        if self.game_over:
            return True


    def draw_key(self, pos):
        if self.maze.key_collected:
            return
        x, y = pos[0] * CELL_SIZE, pos[1] * CELL_SIZE
        cx, cy = x + CELL_SIZE // 2, y + CELL_SIZE // 2
        if self.images.get('key') is not None:
            key_rect = self.images['key'].get_rect(center=(cx, cy))
            self.screen.blit(self.images['key'], key_rect)
        else:
            glow = int(128 + 64 * math.sin(math.radians(self.key_glow)))
            for r in range(25, 15, -2):
                alpha = max(0, glow - (25 - r) * 10)
                s = pygame.Surface((r * 2, r * 2))
                s.set_alpha(alpha); s.fill(KEY_GOLD)
                self.screen.blit(s, s.get_rect(center=(cx, cy)))
            pygame.draw.circle(self.screen, KEY_GOLD, (cx, cy), 12)
            pygame.draw.circle(self.screen, (200, 180, 0), (cx, cy), 12, 2)
            pygame.draw.circle(self.screen, SPACE_BLACK, (cx, cy), 4)
            pygame.draw.rect(self.screen, KEY_GOLD, (cx + 8, cy - 2, 8, 4))

    def draw_character(self, pos, char_type="rebel"):
        x, y = pos[0] * CELL_SIZE, pos[1] * CELL_SIZE
        cx, cy = x + CELL_SIZE // 2, y + CELL_SIZE // 2
        if char_type == "rebel":
            if self.images.get('player'):
                rect = self.images['player'].get_rect(center=(cx, cy))
                self.screen.blit(self.images['player'], rect)
                if self.maze.key_collected:
                    kp = (cx + 20, cy - 20)
                    pygame.draw.circle(self.screen, KEY_GOLD, kp, 8)
                    pygame.draw.circle(self.screen, (200, 180, 0), kp, 8, 2)
                    pygame.draw.circle(self.screen, SPACE_BLACK, kp, 3)
            else:
                pygame.draw.circle(self.screen, REBEL_ORANGE, (cx, cy - 5), 18)
                pygame.draw.circle(self.screen, DARK_GRAY, (cx, cy - 5), 18, 3)
                pygame.draw.ellipse(self.screen, SABER_BLUE, (cx - 12, cy - 12, 24, 10))
                pygame.draw.ellipse(self.screen, REBEL_ORANGE, (cx - 10, cy + 5, 20, 15))
                if self.maze.key_collected:
                    pygame.draw.circle(self.screen, KEY_GOLD, (cx + 15, cy - 15), 6)
        else:  # base
            if self.images.get('base'):
                rect = self.images['base'].get_rect(center=(cx, cy))
                self.screen.blit(self.images['base'], rect)
            else:
                pygame.draw.polygon(self.screen, ICE_BLUE, [(cx, cy-15),(cx-15, cy+10),(cx+15, cy+10)])
                lvl = self.console_font.render("BASE", True, HOTH_WHITE)
                self.screen.blit(lvl, lvl.get_rect(center=(cx, cy)))

    def draw_enemy(self, x, y):
        rect = pygame.Rect(x * CELL_SIZE, y * CELL_SIZE, CELL_SIZE, CELL_SIZE)
        cx, cy = rect.center
        pygame.draw.circle(self.screen, (180, 30, 30), (cx, cy), CELL_SIZE // 3)
        pygame.draw.circle(self.screen, (255, 80, 80), (cx, cy), CELL_SIZE // 3, 2)

    def _algo_viz_colors(self):
        """
        Returns (expansion_overlay_color, path_color) based on current algorithm name.
        """
        name = "" if self.current_algorithm is None else str(self.current_algorithm)
        if "BFS" in name:
            return SABER_BLUE, SABER_BLUE            # blue
        if "DFS" in name:
            return EMPIRE_RED, EMPIRE_RED            # red
        if "A*" in name:
            return A_STAR_YELLOW, A_STAR_YELLOW      # yellow
        if "Greedy" in name:
            return GREEDY_PINK, GREEDY_PINK          # pink
        # default fallback
        return HOLOGRAM_CYAN, HOLOGRAM_CYAN


    def draw_grid(self):
        self.screen.fill(SPACE_BLACK)
        random.seed(42 + self.current_level)
        for _ in range(50 + self.current_level * 8):
            x = random.randint(0, GRID_WIDTH * CELL_SIZE)
            y = random.randint(0, GRID_HEIGHT * CELL_SIZE)
            pygame.draw.circle(self.screen, HOTH_WHITE, (x, y), 1)

        # draw maze area only (actual size is len(grid))
        rows = len(self.maze.grid)
        cols = len(self.maze.grid[0])
        for y in range(rows):
            for x in range(cols):
                r = pygame.Rect(x * CELL_SIZE, y * CELL_SIZE, CELL_SIZE, CELL_SIZE)
                if self.maze.is_wall(x, y):
                    pygame.draw.rect(self.screen, WALL_GRAY, r)
                    pygame.draw.rect(self.screen, ICE_BLUE, r, 2)
                else:
                    pygame.draw.rect(self.screen, HOTH_WHITE, r)
                    pygame.draw.rect(self.screen, ICE_BLUE, r, 1)

        # Search viz (expansion order)
        if not self.manual_mode and self.search.search_order:
            exp_color, _ = self._algo_viz_colors()
            for i in range(min(self.visualization_step, len(self.search.search_order))):
                pos = self.search.search_order[i]
                s = pygame.Surface((CELL_SIZE - 10, CELL_SIZE - 10))
                s.set_alpha(80); s.fill(exp_color)
                self.screen.blit(s, (pos[0]*CELL_SIZE + 5, pos[1]*CELL_SIZE + 5))

        # Solution path (only in AI mode)
        if not self.manual_mode and not self.animating and self.solution_found and self.search.path:
            for pos in self.search.path:
                if pos in (self.maze.start_pos, self.maze.goal_pos):
                    continue
                s = pygame.Surface((CELL_SIZE - 30, CELL_SIZE - 30))
                s.set_alpha(128); s.fill(JEDI_GREEN)  # always green for final path
                self.screen.blit(s, (pos[0]*CELL_SIZE + 15, pos[1]*CELL_SIZE + 15))
            self.drawn_path = True

    def draw_entities(self):
        self.draw_key(self.maze.key_pos)
        self.draw_character(self.maze.goal_pos, "base")
        for e in self.enemies:
            self.draw_enemy(e["x"], e["y"])

        # draw player
        if self.manual_mode or self.autopilot:
            self.draw_character(self.player_pos, "rebel")
        else:
            # AI scan mode: show the rebel where you LAST were in human mode
            pos = self.search.path[0] if self.search.path else self.player_pos
            self.draw_character(pos, "rebel")

    # ---------- UI helpers (wrap & fit) ----------
    def _blit_wrapped(self, text, font, color, x, y, max_width, line_gap=2):
        words = text.split(' ')
        line = ""

        def blit_line(s):
            surf = font.render(s, True, color)
            self.screen.blit(surf, (x, y))
            return y + surf.get_height() + line_gap

        for w in words:
            test = (line + " " + w).strip()
            if font.size(test)[0] <= max_width:
                line = test
            else:
                if line:  # flush current line
                    y = blit_line(line)
                    line = ""

                # word itself too long? hard-wrap by characters
                if font.size(w)[0] > max_width:
                    chunk = ""
                    for ch in w:
                        if font.size(chunk + ch)[0] <= max_width:
                            chunk += ch
                        else:
                            y = blit_line(chunk)
                            chunk = ch
                    line = chunk  # remainder
                else:
                    line = w

        if line:
            y = blit_line(line)
        return y


    def _render_title_fit(self, text, base_size, color, max_width):
        """Return a surface for title text that is scaled down to fit max_width."""
        size = base_size
        while size >= 16:
            f = pygame.font.Font(None, size)
            surf = f.render(text, True, color)
            if surf.get_width() <= max_width:
                return surf
            size -= 2
        # fallback smallest
        return pygame.font.Font(None, 16).render(text, True, color)

    def draw_start_screen(self):
        self.screen.fill(SPACE_BLACK)
        random.seed(42)
        for _ in range(100):
            x = random.randint(0, WINDOW_WIDTH); y = random.randint(0, WINDOW_HEIGHT)
            pygame.draw.circle(self.screen, HOTH_WHITE, (x, y), 1)
        title = pygame.font.Font(None, 48).render("STAR WARS: HOTH ICE MAZE", True, HOLOGRAM_CYAN)
        self.screen.blit(title, title.get_rect(center=(WINDOW_WIDTH//2, 100)))
        subtitle = self.title_font.render("Escape to Echo Base", True, CONSOLE_GREEN)
        self.screen.blit(subtitle, subtitle.get_rect(center=(WINDOW_WIDTH//2, 140)))
        mode = self.title_font.render("SELECT GAME MODE:", True, REBEL_ORANGE)
        self.screen.blit(mode, mode.get_rect(center=(WINDOW_WIDTH//2, 200)))
        for txt, xoff, col1, col2 in (("1 - HUMAN", -200, JEDI_GREEN, CONSOLE_GREEN),
                                      ("2 - AI", 20, EMPIRE_RED, (150,0,0))):
            bg = pygame.Rect(WINDOW_WIDTH//2 + xoff, 250, 180, 80)
            pygame.draw.rect(self.screen, col1, bg); pygame.draw.rect(self.screen, col2, bg, 3)
            t = pygame.font.Font(None, 32).render(txt, True, HOTH_WHITE if "AI" in txt else SPACE_BLACK)
            self.screen.blit(t, t.get_rect(center=(bg.centerx, bg.centery)))
        lines = [
            "MISSION: Collect the golden key, then reach Echo Base",
            "Use WASD or Arrow keys to move (you slide on ice!)",
            "Press 1 for Human Mode  |  Press 2 for AI Mode",
            "Press ESC to quit"
        ]
        for i, L in enumerate(lines):
            color = KEY_GOLD if L.startswith("MISSION") else (HOLOGRAM_CYAN if "Press" in L else HOTH_WHITE)
            t = self.console_font.render(L, True, color)
            self.screen.blit(t, t.get_rect(center=(WINDOW_WIDTH//2, 380 + i*25)))

    def draw_completion_screen(self):
        overlay = pygame.Surface((WINDOW_WIDTH, WINDOW_HEIGHT))
        overlay.set_alpha(180); overlay.fill(SPACE_BLACK); self.screen.blit(overlay, (0,0))
        victory = pygame.font.Font(None, 64).render("MISSION COMPLETE!", True, JEDI_GREEN)
        self.screen.blit(victory, victory.get_rect(center=(WINDOW_WIDTH//2, WINDOW_HEIGHT//2 - 100)))
        success = self.title_font.render("Welcome to Echo Base, Rebel!", True, CONSOLE_GREEN)
        self.screen.blit(success, success.get_rect(center=(WINDOW_WIDTH//2, WINDOW_HEIGHT//2 - 50)))
        for i, opt in enumerate(("R or SPACE - Start New Game", "ESC - Main Menu")):
            t = self.console_font.render(opt, True, HOLOGRAM_CYAN)
            self.screen.blit(t, t.get_rect(center=(WINDOW_WIDTH//2, WINDOW_HEIGHT//2 + i*30)))

    def draw_game_over(self):
        overlay = pygame.Surface((WINDOW_WIDTH, WINDOW_HEIGHT))
        overlay.set_alpha(200); overlay.fill((0,0,0))
        self.screen.blit(overlay, (0,0))
        title = pygame.font.Font(None, 64).render("GAME OVER", True, (255,80,80))
        self.screen.blit(title, title.get_rect(center=(WINDOW_WIDTH//2, WINDOW_HEIGHT//2 - 80)))
        msg = self.title_font.render("A patrol found you!", True, HOTH_WHITE)
        self.screen.blit(msg, msg.get_rect(center=(WINDOW_WIDTH//2, WINDOW_HEIGHT//2 - 40)))
        for i, opt in enumerate(("R - Restart level", "ESC - Main menu")):
            t = self.console_font.render(opt, True, HOLOGRAM_CYAN)
            self.screen.blit(t, t.get_rect(center=(WINDOW_WIDTH//2, WINDOW_HEIGHT//2 + 20 + 28*i)))

    def _current_algo_short(self):
        name = self.current_algorithm or ""
        if "BFS" in name:    return "BFS"
        if "DFS" in name:    return "DFS"
        if "A*" in name:     return "A*"
        if "Greedy" in name: return "Greedy"
        return "—"

    def _draw_mode_badge(self, x, y, w):
        """
        Footer badge showing current mode, heuristic, algorithm, and autopilot status.
        Draws a dark card with a colored pill.
        """
        # card
        h = 90
        card = pygame.Rect(x - 10, y, w, h)
        pygame.draw.rect(self.screen, (20, 25, 35), card, border_radius=12)
        pygame.draw.rect(self.screen, HOLOGRAM_CYAN, card, width=2, border_radius=12)

        # pill
        pill_w, pill_h = 150, 28
        pill = pygame.Rect(card.x + 12, card.y + 10, pill_w, pill_h)
        pill_color = JEDI_GREEN if self.manual_mode else HOLOGRAM_CYAN
        pygame.draw.rect(self.screen, pill_color, pill, border_radius=14)

        # pill label
        pill_text = "HUMAN MODE" if self.manual_mode else "AI MODE"
        label = self.console_font.render(pill_text, True, SPACE_BLACK)
        self.screen.blit(label, label.get_rect(center=pill.center))

        # details (right side)
        info_x = pill.right + 10
        info_y = card.y + 10
        text_area_w = card.right - info_x - 12   # right padding

        # Heuristic (AI only)
        if not self.manual_mode:
            info_y = self._blit_wrapped(
                f"Heuristic: {self.heuristic_used.title()}",
                self.console_font, HOTH_WHITE, info_x, info_y, text_area_w, line_gap=2
            )

        # Short algorithm label + heuristic suffix for A*/Greedy
        short = self._current_algo_short()
        suffix = f" ({self.heuristic_used.title()})" if (short in ("A*", "Greedy") and not self.manual_mode) else ""
        algo_color = REBEL_ORANGE if short != "—" else HOTH_WHITE
        info_y = self._blit_wrapped(
            f"Algorithm: {short}{suffix}",
            self.console_font, algo_color, info_x, info_y, text_area_w, line_gap=2
        )

        # Autopilot
        ap_text  = "ON" if self.autopilot else "OFF"
        ap_color = JEDI_GREEN if self.autopilot else HOTH_WHITE
        self._blit_wrapped(
            f"Autopilot: {ap_text}",
            self.console_font, ap_color, info_x, info_y, text_area_w, line_gap=2
        )


    def draw_ui(self):
        ui_x = GRID_WIDTH * CELL_SIZE + 20
        ui_y = 20
        panel_w = 320
        bg = pygame.Rect(ui_x - 10, ui_y - 10, panel_w, WINDOW_HEIGHT - 40)
        pygame.draw.rect(self.screen, DARK_GRAY, bg); pygame.draw.rect(self.screen, HOLOGRAM_CYAN, bg, 2)

        # Title that always fits the panel
        title_text = f"LEVEL {self.current_level}: {self.maze.get_level_name().upper()}"
        title_surf = self._render_title_fit(title_text, 32, HOLOGRAM_CYAN, max_width=panel_w - 40)
        self.screen.blit(title_surf, (ui_x, ui_y))
        ui_y += title_surf.get_height() + 6

        # Progress
        progress = self.console_font.render(f"PROGRESS: {self.current_level}/{self.max_level}", True, CONSOLE_GREEN)
        self.screen.blit(progress, (ui_x, ui_y))
        ui_y += 35

        # Mode-specific control list
        lines = [
            "MISSION: Collect key, reach base",
            f"KEY: {'COLLECTED' if self.maze.key_collected else 'FIND THE GOLDEN KEY'}",
            "",
            "CONTROLS:",
        ]
        if self.manual_mode:
            lines += [
                "WASD/Arrows - Move pilot",
                "R - Reset current level",
                "M - Toggle Human/AI",
                "ESC - Main menu",
            ]
        else:
            lines += [
                "B - BFS (AI mode)",
                "D - DFS (AI mode)",
                "U - A* (AI mode)",
                "I - Greedy (AI mode)",
                "O - Toggle Heuristic (Manhattan ↔ Euclidean)",
                "P - Benchmark (save to Excel)",
                "A - Autopilot (after scan)",
                "R - Reset current level",
                "M - Toggle Human/AI",
                "ESC - Main menu",
        ]

        if self.current_algorithm and not self.manual_mode:
            short = self._current_algo_short()
            suffix = f" ({self.heuristic_used.title()})" if short in ("A*", "Greedy") else ""
            lines += [
                "",
                f"SCANNER: {short}{suffix}",
                f"SECTORS: {len(self.search.explored)}",
                f"STATUS: {'ROUTE FOUND' if self.solution_found else 'SEARCHING...'}",
            ]

        # Wrap lines to avoid overflow
        max_text_w = panel_w - 40
        for line in lines:
            if line == "":
                ui_y += 8
                continue
            if "MISSION" in line or "LEVEL" in line:
                col = HOLOGRAM_CYAN
            elif "KEY:" in line:
                col = KEY_GOLD if not self.maze.key_collected else JEDI_GREEN
            elif "CONTROLS" in line or "SCANNER" in line:
                col = REBEL_ORANGE
            elif line.startswith("STATUS:"):
                col = CONSOLE_GREEN
            else:
                col = HOTH_WHITE
            ui_y = self._blit_wrapped(line, self.console_font, col, ui_x, ui_y, max_text_w, line_gap=2)

        # footer badge
        footer_y = WINDOW_HEIGHT - 40 - 90  # leave 40px bottom padding and 90px badge height
        self._draw_mode_badge(ui_x, footer_y, panel_w)

        # scanlines effect
        for i in range(0, WINDOW_HEIGHT, 4):
            alpha = 20 if (i + self.scan_lines) % 8 < 4 else 10
            s = pygame.Surface((WINDOW_WIDTH, 1)); s.set_alpha(alpha); s.fill(HOLOGRAM_CYAN)
            self.screen.blit(s, (0, i))

        if self.game_over:
            self.draw_game_over()
        elif self.game_completed:
            self.draw_completion_screen()

    def draw(self):
        if self.show_start_screen:
            self.draw_start_screen()
        else:
            self.draw_grid()
            self.draw_entities()
            self.draw_ui()
        pygame.display.flip()

    # ---------- Loop ----------
    def run(self):
        running = True
        while running:
            running = self.handle_events()
            self.update()
            self.draw()
            self.clock.tick(FPS)
        pygame.quit(); sys.exit()

if __name__ == "__main__":
    game = StarWarsIceMazeGame()
    game.run()


